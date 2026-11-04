using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace FourPrintShops.PrintConfig {
    // Uma cópia do workspace e um DEVMODE integral, exclusivos de uma operação.
    // Este caminho GDI não conhece comandos OEM ou requisitos de Hold.
    internal sealed class WindowsPrintJob : IDisposable {
        private static WindowsPrintJob current;
        private readonly string token=Guid.NewGuid().ToString("N");
        private readonly string sessionId,queueIdentity;
        private readonly int revision;
        private readonly double scale;
        private byte[] bytes,mode;
        private GCHandle pinned;
        private IntPtr document;
        private bool initialized;
        private int pages;
        private WindowsPrintJob(PrinterSession session,double selectedScale) {
            sessionId=session.Id;revision=session.Mode.Revision;queueIdentity=Identity(session);mode=session.Mode.Copy();
            try{scale=EffectiveScale(selectedScale,mode);}catch{Array.Clear(mode,0,mode.Length);throw;}
        }

        [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] private struct DocInfo { public int Size; public string Name,Output,Datatype; public uint Type; }
        [DllImport("kernel32.dll",SetLastError=true)] private static extern bool SetDefaultDllDirectories(uint flags);
        [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] private static extern IntPtr CreateDCW(string driver,string name,string port,IntPtr mode);
        [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] private static extern int StartDocW(IntPtr dc,ref DocInfo info);
        [DllImport("gdi32.dll")] private static extern int StartPage(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern int EndPage(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern int EndDoc(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern int AbortDoc(IntPtr dc);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDF_InitLibrary();
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDF_DestroyLibrary();
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern IntPtr FPDF_LoadMemDocument64(IntPtr bytes,UIntPtr length,string password);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDF_CloseDocument(IntPtr document);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDF_GetPageCount(IntPtr document);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDF_GetFormType(IntPtr document);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern IntPtr FPDF_LoadPage(IntPtr document,int index);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDF_ClosePage(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern double FPDF_GetPageWidth(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern double FPDF_GetPageHeight(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDF_SetPrintMode(int mode);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDF_RenderPage(IntPtr dc,IntPtr page,int x,int y,int width,int height,int rotation,int flags);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFPage_Flatten(IntPtr page,int flags);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFPage_GetAnnotCount(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern IntPtr FPDFPage_GetAnnot(IntPtr page,int index);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFAnnot_GetSubtype(IntPtr annot);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFAnnot_GetFlags(IntPtr annot);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern uint FPDFAnnot_GetAP(IntPtr annot,int mode,IntPtr buffer,uint length);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDFPage_CloseAnnot(IntPtr annot);

        internal static double EffectiveScale(double selected,byte[] mode) {
            if(Double.IsNaN(selected)||Double.IsInfinity(selected)||selected<1||selected>400)throw new InvalidOperationException("Use escala entre 1% e 400%.");
            DevModeState.Validate(mode);
            bool driverScaling=(BitConverter.ToUInt32(mode,72)&16)!=0;
            int driverScale=BitConverter.ToInt16(mode,84);
            // O driver prevalece. Não repetir sua escala no desenho do aplicativo.
            return driverScaling&&driverScale>0&&driverScale!=100?100:selected;
        }
        private static string Identity(PrinterSession session) {
            var info=session.QueueInfo();var text=new StringBuilder(session.Name);
            foreach(string key in new[]{"driver","port","processor","datatype","driverMajorVersion","driverFile","dataFile","configFile"})text.Append('\n').Append(info[key]);
            foreach(string key in new[]{"driverFile","dataFile","configFile"}) {
                string file=Convert.ToString(info[key]);
                using(var stream=File.OpenRead(file))using(var hash=SHA256.Create())text.Append('\n').Append(Convert.ToBase64String(hash.ComputeHash(stream)));
            }
            return text.ToString();
        }
        internal static object Prepare(PrinterSession session,string file,double selectedScale) {
            Clear();var job=new WindowsPrintJob(session,selectedScale);
            try {
                string root=Path.Combine(Path.GetTempPath(),"4PrintShops-print-jobs")+Path.DirectorySeparatorChar;
                file=Path.GetFullPath(file);Guid name;
                if(!file.StartsWith(root,StringComparison.OrdinalIgnoreCase)||!Guid.TryParseExact(Path.GetFileNameWithoutExtension(file),"N",out name)||Path.GetExtension(file)!=".pdf"||(File.GetAttributes(file)&FileAttributes.ReparsePoint)!=0)throw new InvalidOperationException("Cópia temporária inválida.");
                long length=new FileInfo(file).Length;
                if(length<8||length>512L*1024*1024)throw new InvalidOperationException("PDF inválido ou maior que 512 MiB.");
                job.bytes=File.ReadAllBytes(file);
                if(Encoding.ASCII.GetString(job.bytes,0,5)!="%PDF-")throw new InvalidOperationException("PDF inválido.");
                if(!SetDefaultDllDirectories(0x200|0x800))throw new InvalidOperationException("Carregamento seguro de PDFium indisponível.");
                FPDF_InitLibrary();job.initialized=true;
                job.pinned=GCHandle.Alloc(job.bytes,GCHandleType.Pinned);
                job.document=FPDF_LoadMemDocument64(job.pinned.AddrOfPinnedObject(),new UIntPtr((uint)job.bytes.Length),null);
                if(job.document==IntPtr.Zero)throw new InvalidOperationException("PDFium não conseguiu abrir o documento. Use a cópia externa para verificar o arquivo.");
                if(FPDF_GetFormType(job.document)>=2)throw new InvalidOperationException("Formulários XFA ainda exigem o leitor externo para preservar sua aparência.");
                job.pages=FPDF_GetPageCount(job.document);
                if(job.pages<1)throw new InvalidOperationException("O documento não contém páginas imprimíveis.");
                var dimensions=new List<Dictionary<string,object>>();
                // Validar todas as páginas e aparências antes de criar qualquer trabalho.
                for(int index=0;index<job.pages;index++) {
                    IntPtr page=FPDF_LoadPage(job.document,index);
                    if(page==IntPtr.Zero)throw new InvalidOperationException("Não foi possível carregar a página "+(index+1)+".");
                    try {
                        double width=FPDF_GetPageWidth(page),height=FPDF_GetPageHeight(page);
                        Pixels(width,600,job.scale);Pixels(height,600,job.scale);
                        int count=FPDFPage_GetAnnotCount(page);
                        if(count<0)throw new InvalidOperationException("Anotações inválidas.");
                        for(int n=0;n<count;n++) {
                            IntPtr annot=FPDFPage_GetAnnot(page,n);
                            if(annot==IntPtr.Zero)throw new InvalidOperationException("Anotação ilegível.");
                            try {
                                int flags=FPDFAnnot_GetFlags(annot);
                                if(FPDFAnnot_GetSubtype(annot)==20&&(flags&4)!=0&&(flags&3)==0&&FPDFAnnot_GetAP(annot,0,IntPtr.Zero,0)<=2)throw new InvalidOperationException("Campo imprimível sem aparência pronta. Use o leitor externo para preservar o formulário.");
                            } finally {FPDFPage_CloseAnnot(annot);}
                        }
                        if(count>0&&FPDFPage_Flatten(page,1)==0)throw new InvalidOperationException("Não foi possível preservar as aparências de impressão.");
                        dimensions.Add(new Dictionary<string,object>{{"width",width*25.4/72},{"height",height*25.4/72}});
                    } finally {FPDF_ClosePage(page);}
                }
                current=job;
                return new {token=job.token,session=session.Id,revision=session.Mode.Revision,printer=session.Name,pages=job.pages,dimensions=dimensions.Take(50).ToArray(),appScale=job.scale,route="PDFium / Windows GDI"};
            } catch {job.Dispose();throw;}
        }
        internal static int Pixels(double points,int dpi,double percent) {
            double result=points*dpi/72*percent/100;
            if(Double.IsNaN(result)||Double.IsInfinity(result)||result<=0||result>Int32.MaxValue/4)throw new InvalidOperationException("Dimensões da página incompatíveis com o contexto Windows.");
            return Math.Max(1,(int)Math.Round(result));
        }
        internal static object Print(PrinterSession session,string token,string name,string output) {
            var job=current;
            if(job==null||job.token!=token||job.sessionId!=session.Id||job.revision!=session.Mode.Revision)throw new InvalidOperationException("Operação desatualizada. Prepare novamente o documento.");
            current=null; // Consumir antes de qualquer envio. Nunca repetir automaticamente.
            using(job)return job.Run(session,name,output);
        }
        private object Run(PrinterSession session,string name,string output) {
            byte[] confirmed=session.Mode.Copy();IntPtr memory=IntPtr.Zero,dc=IntPtr.Zero;bool started=false;
            try {
                if(!mode.SequenceEqual(confirmed)||queueIdentity!=Identity(session))throw new InvalidOperationException("A fila, o driver ou as propriedades mudaram. Configure novamente.");
                if(!String.IsNullOrEmpty(output)&&!(session.Name=="Microsoft Print to PDF"&&Convert.ToString(session.QueueInfo()["driver"])=="Microsoft Print To PDF"&&Path.IsPathRooted(output)&&Path.GetExtension(output).Equals(".pdf",StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("Destino virtual inválido.");
                memory=Native.Allocate(mode.Length);Marshal.Copy(mode,0,memory,mode.Length);
                dc=CreateDCW("WINSPOOL",session.Name,null,memory);
                if(dc==IntPtr.Zero)throw new InvalidOperationException("O driver não criou um contexto de impressão.");
                var after=Native.ReadMode(memory,mode.Length);
                try{if(!after.SequenceEqual(mode))throw new InvalidOperationException("O driver alterou o DEVMODE confirmado. Envio recusado.");}finally{Array.Clear(after,0,after.Length);}
                int dx=Native.GetDeviceCaps(dc,88),dy=Native.GetDeviceCaps(dc,90),w=Native.GetDeviceCaps(dc,110),h=Native.GetDeviceCaps(dc,111),x=Native.GetDeviceCaps(dc,112),y=Native.GetDeviceCaps(dc,113);
                if(dx<=0||dy<=0||w<=0||h<=0||x<0||y<0||FPDF_SetPrintMode(0)==0)throw new InvalidOperationException("Contexto incompatível com impressão PDFium/GDI.");
                // Conferir a geometria de todas as páginas antes de StartDoc.
                for(int n=0;n<pages;n++) {
                    IntPtr page=FPDF_LoadPage(document,n);
                    if(page==IntPtr.Zero)throw new InvalidOperationException("Página indisponível.");
                    try{Pixels(FPDF_GetPageWidth(page),dx,scale);Pixels(FPDF_GetPageHeight(page),dy,scale);}finally{FPDF_ClosePage(page);}
                }
                if(queueIdentity!=Identity(session))throw new InvalidOperationException("A identidade da fila mudou antes do envio.");
                var info=new DocInfo{Size=Marshal.SizeOf(typeof(DocInfo)),Name=name,Output=String.IsNullOrEmpty(output)?null:output};
                int jobId=StartDocW(dc,ref info);
                if(jobId<=0)throw new InvalidOperationException("O Windows não iniciou o trabalho. Confira a fila antes de tentar novamente.");
                started=true;
                for(int n=0;n<pages;n++) {
                    IntPtr page=FPDF_LoadPage(document,n);
                    if(page==IntPtr.Zero)throw new InvalidOperationException("Falha ao recarregar uma página.");
                    try {
                        int width=Pixels(FPDF_GetPageWidth(page),dx,scale),height=Pixels(FPDF_GetPageHeight(page),dy,scale);
                        if(StartPage(dc)<=0||FPDF_RenderPage(dc,page,(w-width)/2-x,(h-height)/2-y,width,height,0,0x800)==0||EndPage(dc)<=0)throw new InvalidOperationException("Falha durante a geração de uma página. Confira a fila antes de outro trabalho.");
                    } finally {FPDF_ClosePage(page);}
                }
                if(EndDoc(dc)<=0)throw new InvalidOperationException("Conclusão incerta. Confira a fila e a impressora antes de outro trabalho.");
                started=false;
                return new {submitted=true,jobId=jobId,pages=pages,printer=session.Name,appScale=scale,route="PDFium / Windows GDI"};
            } finally {
                if(started&&dc!=IntPtr.Zero)AbortDoc(dc);
                if(dc!=IntPtr.Zero)Native.DeleteDC(dc);
                if(memory!=IntPtr.Zero)Native.Free(memory,mode.Length);
                Array.Clear(confirmed,0,confirmed.Length);
            }
        }
        internal static void Clear(){if(current!=null){current.Dispose();current=null;}}
        public void Dispose() {
            if(document!=IntPtr.Zero){FPDF_CloseDocument(document);document=IntPtr.Zero;}
            if(pinned.IsAllocated)pinned.Free();
            if(initialized){FPDF_DestroyLibrary();initialized=false;}
            if(bytes!=null)Array.Clear(bytes,0,bytes.Length);if(mode!=null)Array.Clear(mode,0,mode.Length);
        }
    }
}
