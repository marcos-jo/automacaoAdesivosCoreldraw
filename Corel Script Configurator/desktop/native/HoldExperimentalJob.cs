using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Globalization;
using FourPrintShops.VirtualPrintLab;

namespace FourPrintShops.PrintConfig {
    internal sealed class HoldExperimentalJob : IDisposable {
        private static HoldExperimentalJob current;
        private static string attemptedSession;
        private readonly string token = Guid.NewGuid().ToString("N");
        private HoldValidation binding;
        private byte[] postscript;

        [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
        private struct RawDocument { public string Name, Output, Datatype; }
        [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] private static extern uint StartDocPrinterW(IntPtr printer,uint level,ref RawDocument info);
        [DllImport("winspool.drv",SetLastError=true)] private static extern bool WritePrinter(IntPtr printer,IntPtr data,uint size,out uint written);
        [DllImport("winspool.drv",SetLastError=true)] private static extern bool EndDocPrinter(IntPtr printer);
        [DllImport("winspool.drv",SetLastError=true)] private static extern bool AbortPrinter(IntPtr printer);

        private static string Digest(byte[] bytes) { using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-","").ToLowerInvariant(); }
        private static KonicaDriverIdentity Identity() {
            IntPtr printer;
            if(!Native.OpenPrinterW(KonicaDriverIdentity.Queue,out printer,IntPtr.Zero))throw new InvalidOperationException("Fila Konica indisponível.");
            try{return KonicaDriverIdentity.Read(printer,KonicaDriverIdentity.Queue);}finally{Native.ClosePrinter(printer);}
        }
        private static Dictionary<string,object> RequireState(PrinterSession session) {
            if(attemptedSession==session.Id)throw new InvalidOperationException("Esta operação já foi utilizada. Inicie deliberadamente uma nova operação.");
            var state=session.VerifyHold();
            var hold=(Dictionary<string,object>)state["hold"];
            var paper=state["geometry"] as Dictionary<string,object>;
            if(session.Name!=KonicaDriverIdentity.Queue||!(bool)hold["holdConfirmed"]||
                !Object.Equals(state["copies"],1)||!Object.Equals(state["driverScale"],100)||paper==null||
                !SupportedPaper(paper)) {
                throw new InvalidOperationException("O ensaio exige Hold, Normal, A4 ou A3 retrato, uma cópia e zoom do driver de 100%.");
            }
            return state;
        }
        private static bool SupportedPaper(Dictionary<string,object> paper) {
            double width=Convert.ToDouble(paper["width"]),height=Convert.ToDouble(paper["height"]);
            return Math.Abs(width-210)<0.3&&Math.Abs(height-297)<0.3 || Math.Abs(width-297)<0.3&&Math.Abs(height-420)<0.3;
        }
        internal static object Prepare(PrinterSession session,string encoded,double scale) {
            var state=RequireState(session);Clear();
            if(Double.IsNaN(scale)||Double.IsInfinity(scale)||scale<1||scale>400)throw new InvalidOperationException("Use escala entre 1% e 400%.");
            var paper=(Dictionary<string,object>)state["geometry"];
            // O contexto informa a folha física. O tamanho lógico público pode ser diferente.
            double paperWidth=Convert.ToDouble(paper["width"])<250?210:297;
            double paperHeight=paperWidth==210?297:420;
            if(encoded.Length>210000)throw new InvalidOperationException("Use o PDF sintético deste laboratório.");
            byte[] bytes=Convert.FromBase64String(encoded);
            string folder=Path.Combine(Path.GetTempPath(),"4PrintShops-hold-experimental",Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string input=Path.Combine(folder,"workspace.pdf"),output=Path.Combine(folder,"captured.ps");
            try {
                File.WriteAllBytes(input,bytes);
                PdfiumVirtualLab.RequireSynthetic(input,Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"synthetic.pdf"));
                byte[] mode=session.Mode.Copy();
                HoldValidation proof=null;
                try {
                    var identity=Identity();
                    proof=new HoldValidation(session.Name,session.Id,session.Mode.Revision,mode,identity.Fingerprint,new HoldDecision{Confirmed=true,Conclusive=true});
                    var prepared=new HoldExperimentalJob();
                    string jobName="4PS HOLD EXP "+prepared.token.Substring(0,8);
                    var capture=PdfiumVirtualLab.Capture(session,input,output,scale,folder,jobName);
                    var ps=File.ReadAllBytes(output);string digest=Digest(ps);
                    string settings=VerifyPostScript(output,paperWidth,paperHeight);
                    if(Digest(File.ReadAllBytes(output))!=digest||!proof.Matches(session.Name,session.Id,session.Mode.Revision,mode,Identity().Fingerprint))throw new InvalidOperationException("A captura mudou durante a verificação.");
                    prepared.binding=proof;prepared.postscript=ps;current=prepared;proof=null;
                    return new {token=current.token,jobName=jobName,printer=session.Name,revision=session.Mode.Revision,paper=paperWidth<250?"A4":"A3",paperWidthMM=paperWidth,paperHeightMM=paperHeight,scale=scale,driverZoom=100,copies=1,destination="Hold",source="PDF sintético A4 do workspace",outputValidated=true,psSha256=digest,cutRisk=(bool)capture["cutRisk"],contentWidthMM=capture["contentWidthMM"],contentHeightMM=capture["contentHeightMM"],printableWidthMM=capture["printableWidthMM"],printableHeightMM=capture["printableHeightMM"],psSettings=settings};
                } finally {Array.Clear(mode,0,mode.Length);if(proof!=null)proof.Dispose();}
            } finally {Array.Clear(bytes,0,bytes.Length);}
        }
        private static string Quote(string value){return "\""+value.Replace("\"","\\\"")+"\"";}
        private static string VerifyPostScript(string file,double paperWidth,double paperHeight) {
            if(!File.Exists(HoldBuildTrust.Ghostscript)||Digest(File.ReadAllBytes(HoldBuildTrust.Ghostscript))!=HoldBuildTrust.GhostscriptSHA256)throw new InvalidOperationException("Ghostscript local de laboratório ausente ou incompatível. Envio bloqueado.");
            string marker=Guid.NewGuid().ToString("N");
            string code="/showpage { ("+marker+" ) print SHKMFeatures /SHKMOutputMethod get == SHKMFeatures /SHKMPrintMode get == SHKMFeatures /Zoom get == SHKMFeatures /HiResPageSize get == currentpagedevice /PageSize get == /#copies where {pop #copies}{1}ifelse == currentpagedevice /NumCopies get dup null eq {pop 1} if == SHKMFeatures /SHKMPdNup get == SHKMFeatures /SHKMPBMode get == (4PS_DETAILS) = [ /SHKMColorMode /SHKMDuplex /SHKMStaple /SHKMInputSlot /PageOrient ] { dup 64 string cvs print (=) print dup SHKMFeatures exch known { SHKMFeatures exch get == } { pop (nao informado) == } ifelse } forall systemdict /showpage get exec } bind def";
            var info=new ProcessStartInfo(HoldBuildTrust.Ghostscript,"-dSAFER -dBATCH -dNOPAUSE -sDEVICE=nullpage -c "+Quote(code)+" -f "+Quote(file)) {UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true};
            using(var process=Process.Start(info)) {
                var output=process.StandardOutput.ReadToEndAsync();var errors=process.StandardError.ReadToEndAsync();
                if(!process.WaitForExit(30000)){process.Kill();throw new InvalidOperationException("Interpretação PS sem resposta. Envio bloqueado.");}
                string text=output.GetAwaiter().GetResult().Replace("\r","");errors.GetAwaiter().GetResult();
                string sizes=String.Format(CultureInfo.InvariantCulture,"[{0:0} {1:0}]\n[{2:0} {3:0}]\n",paperWidth*1000,paperHeight*1000,Math.Round(paperWidth*72/25.4),Math.Round(paperHeight*72/25.4));
                string expected=marker+" /Hold\n/Print\n100\n"+sizes+"1\n1\n/None\n/None\n";
                if(process.ExitCode!=0||text.Split(new[]{marker},StringSplitOptions.None).Length!=2||!text.Contains(expected))throw new InvalidOperationException("O PS executado não comprovou Hold, papel do driver, zoom 100%, uma cópia e ausência de imposição. Envio bloqueado.");
                int start=text.IndexOf("4PS_DETAILS\n",StringComparison.Ordinal);
                if(start<0)throw new InvalidOperationException("Diagnóstico do PS inconclusivo.");
                return String.Join("\n",text.Substring(start+12).Split('\n').Take(5).ToArray()).Trim();
            }
        }
        internal static object Send(PrinterSession session,string token) {
            RequireState(session);
            if(current==null||current.token!=token)throw new InvalidOperationException("Captura aprovada indisponível.");
            IntPtr printer=IntPtr.Zero,memory=IntPtr.Zero;bool started=false;uint job=0;
            string name="4PS HOLD EXP "+current.token.Substring(0,8);
            try {
                if(!Native.OpenPrinterW(KonicaDriverIdentity.Queue,out printer,IntPtr.Zero))throw new InvalidOperationException();
                byte[] mode=session.Mode.Copy();
                try {
                    var actual=KonicaDriverIdentity.Read(printer,session.Name);
                    if(!current.binding.Matches(session.Name,session.Id,session.Mode.Revision,mode,actual.Fingerprint))throw new InvalidOperationException("A configuração mudou após a captura.");
                } finally {Array.Clear(mode,0,mode.Length);}
                attemptedSession=session.Id; // Consumir somente esta operação antes de abrir o trabalho.
                var info=new RawDocument{Name=name,Datatype="RAW"};
                job=StartDocPrinterW(printer,1,ref info);
                if(job==0)throw new InvalidOperationException();started=true;
                memory=Native.Allocate(65536);
                for(int offset=0;offset<current.postscript.Length;) {
                    int count=Math.Min(65536,current.postscript.Length-offset);Marshal.Copy(current.postscript,offset,memory,count);
                    uint written;if(!WritePrinter(printer,memory,(uint)count,out written)||written==0||written>count)throw new InvalidOperationException();
                    offset+=(int)written;
                }
                if(!EndDocPrinter(printer))throw new InvalidOperationException();started=false;
                return new {submitted=true,jobId=job,jobName=name,destination="Hold",message="Trabalho entregue ao spooler. Confira a retenção na Konica antes de qualquer liberação física."};
            } catch {throw new InvalidOperationException("Resultado de envio incerto. Confira a fila e a Konica antes de iniciar outro ensaio. Não reenviar automaticamente.");}
            finally {if(started&&printer!=IntPtr.Zero)AbortPrinter(printer);if(printer!=IntPtr.Zero)Native.ClosePrinter(printer);Native.Free(memory,65536);Clear();}
        }
        internal static void Clear(){if(current!=null)current.Dispose();current=null;}
        public void Dispose(){if(binding!=null)binding.Dispose();if(postscript!=null)Array.Clear(postscript,0,postscript.Length);postscript=null;}
    }
}
