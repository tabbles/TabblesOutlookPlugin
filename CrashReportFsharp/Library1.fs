module CrashReportFsharp
open System;
open System.Runtime.InteropServices
open System.Threading.Tasks;
open System.IO.Pipes;
open System.Reflection;
open System.ServiceModel;
//[<assembly:AssemblyKeyFileAttribute ("Yellow-Blue-soft.snk.pfx")>]
//do();


let emptyStringOfNull s =
        if s = null then "" else s

let debugFsharpIsWorking () =
        "fsharp is working"

let ifNull2 valueIfNull (nullable ) =
        if nullable = null then
                valueIfNull
        else
                nullable


let nl = Environment.NewLine + Environment.NewLine;

let rec create_message_of_exception (ex: Exception) msg  level = 
        
        let hres = Marshal.GetHRForException(ex).ToString()
        let tab = "                    ";
        let nl = Environment.NewLine + Environment.NewLine
        let formatStackTrace (stackt : string) =
                try
                        let lines = 
                        
                                stackt.Split([| Environment.NewLine |], StringSplitOptions.None);
                        let l2 = [ for l in lines do
                                        yield tab + l + nl ]

                        List.fold  (fun a b -> a + b) "" l2
                with
                | :? NullReferenceException ->
                        ""
        
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
                let msg' = 
                
                        let stackTr =
                                try
                                        ex.StackTrace |> ifNull2 ""
                                with
                                | e ->  "stacktrace cannot be obtained! crashed: " + e.GetType().ToString() + " --- " +  e.Message 

                        msg + nl +
                                tab + "-------------------- Level " + string level + " -------------" + nl +  
                                tab + "Exception type = " + ex.GetType().ToString() + nl +
                                tab +  "hresult = " + hres  + nl +
                                tab +  sqlDetail +  "Message = " + ex.Message + nl +
                                tab + "StackTrace = " + nl +
                                        (formatStackTrace stackTr)
                create_message_of_exception  ex.InnerException   msg'  (level + 1)


let stringOfException (e: Exception) = 
        try
                create_message_of_exception e "" 0       
        with
        | er -> "error in string-of-exception: " + er.GetType().ToString() + " --- " + er.Message


let wrapNeverCrash logError f =
        try

                f()
        with
        | e -> 
                try
                        logError ("wrapNeverCrash : internal error in " + f.ToString() + " : "  + stringOfException e)
                                
                with
                | _ -> ()

//let preamboloCrashLogAutomatico () =
//        let windowsUser = 
//                        Environment.UserName |> ifNull2 ""
//        let machineName = 
//                Environment.MachineName |> ifNull2 ""
//        let emailStr = 
//                       "win user = " + windowsUser + "; machineName = " + machineName  + Environment.NewLine + Environment.NewLine 
//        
//        emailStr 



/// <summary>
/// 
/// </summary>
/// <param name="logError">La funzione che stampa nel log. serve ad accorgersi di errori interni di execInThreadForceNewThreadDur</param>
/// <param name="f"></param>
let execInThreadForceNewThreadDur longRunning (logError: string -> unit) (f) =
        
        let fNeverCrash()  = wrapNeverCrash  logError f
        let t = Task.Factory.StartNew( fNeverCrash , if longRunning then  Threading.Tasks.TaskCreationOptions.LongRunning else Threading.Tasks.TaskCreationOptions.None )
        ()

let g_lock = "foo2"

let critSec2 f = 
        lock g_lock f

//let critSecAction (f: Action) = 
//        lock g_lock (fun () -> f.Invoke())
//
//let critSecFunc (f: Func<'a>) = 
//        lock g_lock (fun () -> f.Invoke())



//let addressWebService = "http://tag-fs.com.iis3004.databasemart.net/updateservice/TabblesServ.svc";

//let addressWebService = "http://register.tag-forge.com/updateservice/TabblesServ.svc";

                        

/// <summary>
/// 
/// </summary>
/// <param name="logError">funzione che stampa nel log una stringa, utile in caso fallisca la stessa sendSilentCrash</param>
/// <param name="testo"></param>

        
let scramble_string2 (s: string) = 
        let by =  Text.Encoding.UTF8.GetBytes(s)
        let by = 
                let version  = 1uy
                Array.append  [| version |] by
        let shift_byte (b : byte) = 
                let i = int b
                (i + 55) % 255 |> byte
        let by = by |> Array.map shift_byte 
        Convert.ToBase64String by
         
let nuOf (x: 'a) = 
        new Nullable<'a>(x)


let noAction () = ()



type createPipeServerResult = CpsrOkFirstTry of NamedPipeServerStream  
                                | CpsrOkSecondTry of NamedPipeServerStream * string 
                                | CpsrOkThirdTry of NamedPipeServerStream * string  
                                | CpsrFailed of string
                                | CpsrInternalErrorAfterServerCreation of Exception




exception TestPipeException

/// <summary>
/// Creates a pipe server in 3 different ways. If one or more ways fails, it sends a silent crash report.
/// </summary>
/// <param name="logError"></param>
/// <param name="pipeName"></param>
let createPipeServerSafeForExternalProcesses preambolo (logError: string -> unit) pipeName =
        try

                let ps = new PipeSecurity();
                
                                        
                ps.AddAccessRule(new PipeAccessRule(new Security.Principal.SecurityIdentifier(Security.Principal.WellKnownSidType.WorldSid, null), 
                                                        PipeAccessRights.FullControl, 
                                                        System.Security.AccessControl.AccessControlType.Allow));


                let serv = new IO.Pipes.NamedPipeServerStream(
                                pipeName, 
                                IO.Pipes.PipeDirection.InOut, // importante! non solo in, altrimenti crasha, e dice UnauthorizedAccessException, fuorviante, perché non è pipesecurity! vedi blog: http://adventuresindevelopment.blogspot.it/2008/07/named-pipes-issue-systemunauthorizedacc.html
                                10,
                                PipeTransmissionMode.Message,   // Message-based communication
                                PipeOptions.None,
                                1024,
                                1024,
                                ps);
                CpsrOkFirstTry serv
        with
        | e1 ->

                try
                        let ps = new PipeSecurity();
                                        

                        ps.AddAccessRule(new PipeAccessRule(new Security.Principal.SecurityIdentifier(Security.Principal.WellKnownSidType.BuiltinUsersSid, null), 
                                                                PipeAccessRights.FullControl, 
                                                                System.Security.AccessControl.AccessControlType.Allow));
                        ps.AddAccessRule(new PipeAccessRule(new Security.Principal.SecurityIdentifier(Security.Principal.WellKnownSidType.CreatorOwnerSid, null), 
                                                                        PipeAccessRights.FullControl, 
                                                                        System.Security.AccessControl.AccessControlType.Allow));
                        ps.AddAccessRule(new PipeAccessRule(new Security.Principal.SecurityIdentifier(Security.Principal.WellKnownSidType.LocalSystemSid, null), 
                                                                        PipeAccessRights.FullControl, 
                                                                        System.Security.AccessControl.AccessControlType.Allow));


                        let serv = new IO.Pipes.NamedPipeServerStream(
                                        pipeName, 
                                        IO.Pipes.PipeDirection.InOut, // importante! non solo in, altrimenti crasha, e dice UnauthorizedAccessException, fuorviante, perché non è pipesecurity! vedi blog: http://adventuresindevelopment.blogspot.it/2008/07/named-pipes-issue-systemunauthorizedacc.html
                                        10,
                                        PipeTransmissionMode.Message,   // Message-based communication
                                        PipeOptions.None,
                                        1024,
                                        1024,
                                        ps);

                        try
                                let strCrashId = "Failed attempt to create pipe with pipesecurity 1. second attempt worked." 
                                let stackTrace =  preambolo + nl  + stringOfException e1
                                let windowsUser = Environment.UserName |> emptyStringOfNull
                                let machineName = Environment.MachineName |> emptyStringOfNull
                                
                                let strExc = strCrashId + nl + stackTrace
                                logError (strExc);
                        
                                CpsrOkSecondTry ( serv, strExc)
                        with
                        | e ->
                                serv.Dispose();
                                CpsrInternalErrorAfterServerCreation e

                with
                | e2 ->
                        try
                                let pipeServer =                                         
                                                new IO.Pipes.NamedPipeServerStream(
                                                                        maxNumberOfServerInstances = 10,
                                                                        transmissionMode = PipeTransmissionMode.Message,   // Message-based communication
                                                                        options = PipeOptions.None, 
                                                                        pipeName = pipeName, 
                                                                        direction = IO.Pipes.PipeDirection.InOut // importante! non solo in, altrimenti crasha, e dice UnauthorizedAccessExcepttion, fuorviante, perché non è pipesecirity! vedi blog: http://adventuresindevelopment.blogspot.it/2008/07/named-pipes-issue-systemunauthorizedacc.html
                                                                        )       

                                try
                                        let strCrashId = "Failed 2nd attempt to create pipe with pipesecurity. third attempt worked." 
                                        let stackTrace = preambolo + nl + "error 1: " + stringOfException e1 + nl
                                                                +  "error 2: " + stringOfException e2 + nl
                                        let windowsUser = Environment.UserName |> emptyStringOfNull
                                        let machineName = Environment.MachineName |> emptyStringOfNull
                                        
                                        let strExc = strCrashId + nl + stackTrace
                                        logError (strExc);
                                        CpsrOkThirdTry (pipeServer , strExc);
                                with
                                | e ->
                                        pipeServer.Dispose();
                                        CpsrInternalErrorAfterServerCreation e  
                        with
                        | e3 ->
                                let strCrashId = "Failed all attempts to create pipe: " 
                                let stackTrace = preambolo + nl + "error 1 : " + stringOfException e1 + nl 
                                                 + "error 2 : " + stringOfException e2 + nl 
                                                 + "error 3 : " + stringOfException e3 + nl 
                                let windowsUser = Environment.UserName |> emptyStringOfNull
                                let machineName = Environment.MachineName |> emptyStringOfNull
                                
                                let strExc = strCrashId + nl + stackTrace
                                logError (strExc);
                                CpsrFailed strExc

                        
/// <summary>
/// Use when you need to compute a value, and if the computation crashes fallback to a default value.
/// </summary>
/// <param name="logError"></param>
/// <param name="preambolo"></param>
/// <param name="valueIfThrows"></param>
/// <param name="f"></param>
//let tryCatchCompute logError preambolo valueIfThrows f =
//        try
//                f ()
//        with
//        | e ->
//                let str = stringOfException e;
//                logError str;
//                sendSilentCrashIfEnoughTimePassed2 logError (preambolo + nl + str);
//                valueIfThrows
                                                