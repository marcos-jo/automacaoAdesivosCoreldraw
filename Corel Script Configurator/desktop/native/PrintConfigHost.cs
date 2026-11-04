using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace FourPrintShops.PrintConfig {
    internal static class Host {
        [STAThread] private static int Main(string[] args) {
            int parent;
            if(args.Length!=1||!Int32.TryParse(args[0],out parent)||parent<=0)return 2;
            Console.InputEncoding=new UTF8Encoding(false);Console.OutputEncoding=new UTF8Encoding(false);
            var monitor=new Thread(delegate(){for(;;){try{using(var process=Process.GetProcessById(parent)){if(process.HasExited)Environment.Exit(0);}}catch{Environment.Exit(0);}Thread.Sleep(1000);}});monitor.IsBackground=true;monitor.Start();
            var json=new JavaScriptSerializer {MaxJsonLength=262144,RecursionLimit=12};PrinterSession session=null;
            try { string line;
                while((line=Console.ReadLine())!=null){int id=0;try{
#if HOLD_EXPERIMENTAL
                    if(line.Length>230000)throw new InvalidOperationException("Solicitação inválida.");
#else
                    if(line.Length>4096)throw new InvalidOperationException("Solicitação inválida.");
#endif
                    var request=json.Deserialize<Dictionary<string,object>>(line);id=Convert.ToInt32(request["id"]);string action=Convert.ToString(request["action"]);object result;
                    if(action=="list")result=Native.Printers();
                    else if(action=="select"){
#if WINDOWS_PRINT
                        WindowsPrintJob.Clear();
#endif
                        string name=Convert.ToString(request["printer"]);bool installed=false;
                        foreach(var item in Native.Printers())if(String.Equals((string)item["name"],name,StringComparison.Ordinal))installed=true;
                        if(!installed)throw new InvalidOperationException("Selecione uma impressora instalada.");
                        var replacement=new PrinterSession(name);
                        try{result=replacement.Summary();}catch{replacement.Dispose();throw;}
                        if(session!=null)session.Dispose();session=replacement;
                    }else if(action=="properties"||action=="summary"||action=="verify-hold"){
                        if(session==null||Convert.ToString(request["session"])!=session.Id||Convert.ToInt32(request["revision"])!=session.Mode.Revision)throw new InvalidOperationException("Sessão desatualizada. Selecione novamente a impressora.");
                        if(action=="verify-hold")result=session.VerifyHold();
                        else if(action=="summary")result=session.Summary();else{
                            long handle=Int64.Parse(Convert.ToString(request["owner"]));uint ownerProcess;
                            if(handle!=0&&(Native.GetWindowThreadProcessId(new IntPtr(handle),out ownerProcess)==0||ownerProcess!=parent))throw new InvalidOperationException("Janela proprietária inválida.");
#if WINDOWS_PRINT
                            WindowsPrintJob.Clear();
#endif
                            result=session.Properties(new IntPtr(handle),session.Mode.Revision);
                        }
                    }
#if HOLD_EXPERIMENTAL
                    else if(action=="prepare-hold"||action=="send-hold"){
#if WINDOWS_PRINT
                        WindowsPrintJob.Clear();
#endif
                        if(session==null||Convert.ToString(request["session"])!=session.Id||Convert.ToInt32(request["revision"])!=session.Mode.Revision)throw new InvalidOperationException("Sessão desatualizada.");
                        result=action=="prepare-hold"?HoldExperimentalJob.Prepare(session,Convert.ToString(request["pdf"]),request.ContainsKey("scale")?Convert.ToDouble(request["scale"]):100):HoldExperimentalJob.Send(session,Convert.ToString(request["token"]));
                    }
#endif
#if WINDOWS_PRINT
                    else if(action=="prepare-document"||action=="print-document"||action=="cancel-document"){
                        if(session==null||Convert.ToString(request["session"])!=session.Id||Convert.ToInt32(request["revision"])!=session.Mode.Revision)throw new InvalidOperationException("Sessão de impressão desatualizada.");
                        if(action=="prepare-document"){
                            HoldExperimentalJob.Clear();
                            result=WindowsPrintJob.Prepare(session,Convert.ToString(request["file"]),Convert.ToDouble(request["scale"]));
                        }else if(action=="print-document")result=WindowsPrintJob.Print(session,Convert.ToString(request["token"]),Convert.ToString(request["name"]),request.ContainsKey("output")?Convert.ToString(request["output"]):null);
                        else {WindowsPrintJob.Clear();result=true;}
                    }
#endif
                    else if(action=="release"){
#if WINDOWS_PRINT
                        WindowsPrintJob.Clear();
#endif
#if HOLD_EXPERIMENTAL
                        HoldExperimentalJob.Clear();
#endif
                        if(session!=null)session.Dispose();session=null;result=true;}
                    else throw new InvalidOperationException("Comando não permitido.");
                    Console.WriteLine(json.Serialize(new {id=id,ok=true,result=result}));
                }catch(InvalidOperationException error){Console.WriteLine(json.Serialize(new {id=id,ok=false,error=error.Message}));}
                catch{Console.WriteLine(json.Serialize(new {id=id,ok=false,error="Falha no auxiliar nativo. Se a impressão foi iniciada, confira a fila antes de outro trabalho."}));}
                }
            }finally{if(session!=null)session.Dispose();
#if WINDOWS_PRINT
                WindowsPrintJob.Clear();
#endif
#if HOLD_EXPERIMENTAL
                HoldExperimentalJob.Clear();
#endif
            }
            return 0;
        }
    }
}
