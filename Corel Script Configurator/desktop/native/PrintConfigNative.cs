using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace FourPrintShops.PrintConfig {
    internal static class Native {
        [StructLayout(LayoutKind.Sequential)] internal struct Printer4 { public IntPtr Name, Server; public uint Attributes; }
        [StructLayout(LayoutKind.Sequential)] internal struct Printer2 {
            public IntPtr Server,Name,Share,Port,Driver,Comment,Location,DevMode,SepFile,Processor,Datatype,Parameters,Security;
            public uint Attributes,Priority,DefaultPriority,StartTime,UntilTime,Status,Jobs,AveragePPM;
        }
        [StructLayout(LayoutKind.Sequential)] internal struct Driver2 { public uint Version; public IntPtr Name,Environment,Driver,Data,Config; }
        [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] internal static extern bool GetPrinterDriverW(IntPtr printer,string environment,uint level,IntPtr data,uint size,out uint needed);
        [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] internal static extern bool EnumPrintersW(uint flags,string name,uint level,IntPtr data,uint size,out uint needed,out uint count);
        [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] internal static extern bool OpenPrinterW(string name,out IntPtr printer,IntPtr defaults);
        [DllImport("winspool.drv",SetLastError=true)] internal static extern bool ClosePrinter(IntPtr printer);
        [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] internal static extern int DocumentPropertiesW(IntPtr owner,IntPtr printer,string name,IntPtr output,IntPtr input,uint flags);
        [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] internal static extern bool GetPrinterW(IntPtr printer,uint level,IntPtr data,uint size,out uint needed);
        [DllImport("winspool.drv",CharSet=CharSet.Unicode)] internal static extern int DeviceCapabilitiesW(string name,string port,ushort capability,IntPtr output,IntPtr mode);
        [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] internal static extern IntPtr CreateICW(string driver,string name,string port,IntPtr mode);
        [DllImport("gdi32.dll")] internal static extern int GetDeviceCaps(IntPtr context,int index);
        [DllImport("gdi32.dll")] internal static extern bool DeleteDC(IntPtr context);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        internal static IntPtr Allocate(int size) { if(size<=0||size>4*1024*1024)throw new InvalidOperationException("Tamanho inválido retornado pelo driver."); IntPtr p=Marshal.AllocHGlobal(size); for(int n=0;n<size;n++)Marshal.WriteByte(p,n,0);return p; }
        internal static void Free(IntPtr p,int size) { if(p==IntPtr.Zero)return; for(int n=0;n<size;n++)Marshal.WriteByte(p,n,0);Marshal.FreeHGlobal(p); }
        internal static string Text(IntPtr p) { return p==IntPtr.Zero?null:Marshal.PtrToStringUni(p); }
        internal static byte[] ReadMode(IntPtr p,int capacity) {
            if(p==IntPtr.Zero||capacity<102)throw new InvalidOperationException("DEVMODE ausente.");
            int size=(ushort)Marshal.ReadInt16(p,68),extra=(ushort)Marshal.ReadInt16(p,70),total=size+extra;
            if(size<102||total>capacity)throw new InvalidOperationException("DEVMODE incompleto.");
            byte[] result=new byte[total];Marshal.Copy(p,result,0,total);DevModeState.Validate(result);return result;
        }
        internal static List<Dictionary<string,object>> Printers() {
            uint needed,count;EnumPrintersW(6,null,4,IntPtr.Zero,0,out needed,out count);
            var result=new List<Dictionary<string,object>>();if(needed==0){if(Marshal.GetLastWin32Error()!=0)throw new InvalidOperationException("Não foi possível listar impressoras.");return result;}
            int capacity=checked((int)needed);IntPtr p=Allocate(capacity);
            try { if(!EnumPrintersW(6,null,4,p,needed,out needed,out count))throw new InvalidOperationException("Não foi possível listar impressoras.");
                int stride=Marshal.SizeOf(typeof(Printer4));if(count>1024||count*stride>needed)throw new InvalidOperationException("Lista de impressoras inválida.");
                for(int n=0;n<count;n++){var printer=(Printer4)Marshal.PtrToStructure(IntPtr.Add(p,n*stride),typeof(Printer4));string name=Text(printer.Name);
                    if(!String.IsNullOrWhiteSpace(name)&&name.Length<=512)result.Add(new Dictionary<string,object>{{"name",name}});}
                return result;
            } finally {Free(p,capacity);}
        }
    }

    internal sealed class PrinterSession : IDisposable {
        internal readonly string Name, Id;
        internal readonly DevModeState Mode;
        private IntPtr printer;
        private HoldValidation holdValidation;
        internal PrinterSession(string name) {
            Name=name;Id=Guid.NewGuid().ToString("N");
            if(!Native.OpenPrinterW(name,out printer,IntPtr.Zero))throw new InvalidOperationException("Não foi possível acessar a impressora selecionada.");
            byte[] initial=null;
            try { initial=Initial();Mode=new DevModeState(initial); }
            catch { Dispose();throw; }
            finally { if(initial!=null)Array.Clear(initial,0,initial.Length); }
        }
        private byte[] Initial() {
            int capacity=Native.DocumentPropertiesW(IntPtr.Zero,printer,Name,IntPtr.Zero,IntPtr.Zero,0);
            IntPtr output=Native.Allocate(capacity),input=IntPtr.Zero,info=IntPtr.Zero;int infoSize=0;
            try {
                uint needed;Native.GetPrinterW(printer,9,IntPtr.Zero,0,out needed);
                if(needed>0&&needed<=4*1024*1024){infoSize=(int)needed;info=Native.Allocate(infoSize);
                    if(Native.GetPrinterW(printer,9,info,needed,out needed)){
                        IntPtr stored=Marshal.ReadIntPtr(info);long offset=stored.ToInt64()-info.ToInt64();
                        if(stored!=IntPtr.Zero&&offset>=0&&offset+102<=infoSize){byte[] user=Native.ReadMode(stored,infoSize-(int)offset);
                            try { input=Native.Allocate(capacity);if(user.Length>capacity)throw new InvalidOperationException("Configuração do usuário incompatível.");Marshal.Copy(user,0,input,user.Length); }
                            finally {Array.Clear(user,0,user.Length);}
                        }
                    }
                }
                uint flags=2u|(input==IntPtr.Zero?0u:8u); // OUT_BUFFER e, quando presente, IN_BUFFER. Nunca persistir.
                if(Native.DocumentPropertiesW(IntPtr.Zero,printer,Name,output,input,flags)!=1)throw new InvalidOperationException("O driver não forneceu uma configuração válida.");
                return Native.ReadMode(output,capacity);
            } finally {Native.Free(output,capacity);Native.Free(input,capacity);Native.Free(info,infoSize);}
        }
        internal Dictionary<string,object> Properties(IntPtr owner,int expectedRevision) {
            if(expectedRevision!=Mode.Revision)throw new InvalidOperationException("Configuração desatualizada. Selecione novamente a impressora.");
            byte[] current=Mode.Copy(),candidate=null;int capacity=Native.DocumentPropertiesW(owner,printer,Name,IntPtr.Zero,IntPtr.Zero,0);
            IntPtr input=IntPtr.Zero,output=IntPtr.Zero;
            try {
                if(current.Length>capacity)throw new InvalidOperationException("O driver mudou. Selecione novamente a impressora.");
                input=Native.Allocate(capacity);output=Native.Allocate(capacity);Marshal.Copy(current,0,input,current.Length);
                int answer=Native.DocumentPropertiesW(owner,printer,Name,output,input,2|4|8);
                if(answer==2){var canceled=Summary();canceled["canceled"]=true;return canceled;}
                if(answer!=1)throw new InvalidOperationException("O driver não confirmou as propriedades.");
                candidate=Native.ReadMode(output,capacity);
                var summary=Describe(candidate); // Só aceitar depois de validar o resumo; falha conserva a versão anterior.
                Mode.Accept(candidate,expectedRevision);
                if(holdValidation!=null)holdValidation.Dispose();
                holdValidation=null;
                AddIdentity(summary);summary["canceled"]=false;
                summary["hold"]=HoldValidation.Unverified("Confirme novamente o modo Hold após alterar as propriedades.");
                return summary;
            } finally {Array.Clear(current,0,current.Length);if(candidate!=null)Array.Clear(candidate,0,candidate.Length);Native.Free(input,capacity);Native.Free(output,capacity);}
        }
        internal Dictionary<string,object> QueueInfo() {
            uint needed;Native.GetPrinterW(printer,2,IntPtr.Zero,0,out needed);
            int size=checked((int)needed);IntPtr data=Native.Allocate(size);
            var result=new Dictionary<string,object>();
            try {
                if(!Native.GetPrinterW(printer,2,data,(uint)size,out needed))throw new InvalidOperationException("A fila não está disponível.");
                var info=(Native.Printer2)Marshal.PtrToStructure(data,typeof(Native.Printer2));
                result["driver"]=Native.Text(info.Driver);result["port"]=Native.Text(info.Port);
                result["processor"]=Native.Text(info.Processor);result["datatype"]=Native.Text(info.Datatype);
            } finally {Native.Free(data,size);}
            Native.GetPrinterDriverW(printer,null,2,IntPtr.Zero,0,out needed);size=checked((int)needed);data=Native.Allocate(size);
            try {
                if(!Native.GetPrinterDriverW(printer,null,2,data,(uint)size,out needed))throw new InvalidOperationException("Não foi possível consultar o driver.");
                var info=(Native.Driver2)Marshal.PtrToStructure(data,typeof(Native.Driver2));
                result["driverMajorVersion"]=info.Version;result["driverFile"]=Native.Text(info.Driver);
                result["dataFile"]=Native.Text(info.Data);result["configFile"]=Native.Text(info.Config);
                string core=System.IO.Path.GetFileName(Native.Text(info.Driver))??"";
                result["language"]=core.Equals("PSCRIPT5.DLL",StringComparison.OrdinalIgnoreCase)?"PostScript (driver Windows)":"Linguagem produzida pelo driver; PDL não exposta";
            } finally {Native.Free(data,size);}
            return result;
        }
        private void AddIdentity(Dictionary<string,object> result) {result["session"]=Id;result["revision"]=Mode.Revision;result["printer"]=Name;result["queue"]=QueueInfo();}
        internal Dictionary<string,object> Summary() {byte[] bytes=Mode.Copy();try{var result=Describe(bytes);AddIdentity(result);result["hold"]=CurrentHold(bytes);return result;}finally{Array.Clear(bytes,0,bytes.Length);}}
        internal Dictionary<string,object> VerifyHold() {
            byte[] bytes=Mode.Copy();
            try {
                if(holdValidation!=null)holdValidation.Dispose();
                holdValidation=null;
                var before=KonicaDriverIdentity.Read(printer,Name);
                var decision=HoldTicketReader.Read(Name,bytes,before.TicketNamespace);
                var after=KonicaDriverIdentity.Read(printer,Name);
                if(before.Fingerprint!=after.Fingerprint)throw new InvalidOperationException();
                holdValidation=new HoldValidation(Name,Id,Mode.Revision,bytes,before.Fingerprint,decision);
            } catch {
                var failed=Summary();
                failed["hold"]=HoldValidation.Unverified("Envio bloqueado. Não foi possível validar o PrintTicket e a identidade do driver Konica.");
                return failed;
            } finally {Array.Clear(bytes,0,bytes.Length);}
            return Summary();
        }
        private Dictionary<string,object> CurrentHold(byte[] bytes) {
            if(holdValidation==null)return HoldValidation.Unverified("Selecione Hold nas propriedades e verifique esta sessão.");
            try {
                var identity=KonicaDriverIdentity.Read(printer,Name);
                if(holdValidation.Matches(Name,Id,Mode.Revision,bytes,identity.Fingerprint))return holdValidation.Describe();
            } catch {}
            holdValidation.Dispose();holdValidation=null;
            return HoldValidation.Unverified("A configuração ou o driver mudou. Verifique novamente o modo Hold.");
        }
        private static object Field(byte[] bytes,uint mask,int offset) {return (BitConverter.ToUInt32(bytes,72)&mask)!=0?(object)(int)BitConverter.ToInt16(bytes,offset):null;}
        private Dictionary<string,object> Describe(byte[] bytes) {
            DevModeState.Validate(bytes);IntPtr mode=Native.Allocate(bytes.Length),context=IntPtr.Zero;
            try {
                Marshal.Copy(bytes,0,mode,bytes.Length);
                object orientation=Field(bytes,1,76),paperId=Field(bytes,2,78),copies=Field(bytes,256,86),tray=Field(bytes,512,88),duplex=Field(bytes,4096,94),scale=Field(bytes,16,84),collate=Field(bytes,32768,100);
                var result=new Dictionary<string,object>{{"orientation",orientation},{"copies",copies},{"duplex",duplex},{"driverScale",scale},{"collate",collate},{"trayId",tray},{"paperId",paperId}};
                result["paper"]=Paper(mode,paperId,orientation,bytes);result["tray"]=CapabilityName(mode,tray,6,12,24);
                context=Native.CreateICW("WINSPOOL",Name,null,mode);
                if(context==IntPtr.Zero){result["geometry"]=null;return result;}
                int dx=Native.GetDeviceCaps(context,88),dy=Native.GetDeviceCaps(context,90),w=Native.GetDeviceCaps(context,110),h=Native.GetDeviceCaps(context,111),x=Native.GetDeviceCaps(context,112),y=Native.GetDeviceCaps(context,113),pw=Native.GetDeviceCaps(context,8),ph=Native.GetDeviceCaps(context,10);
                if(dx<=0||dy<=0||w<=0||h<=0||pw<=0||ph<=0||x<0||y<0||x+pw>w+2||y+ph>h+2){result["geometry"]=null;return result;}
                result["geometry"]=new Dictionary<string,object>{{"width",w*25.4/dx},{"height",h*25.4/dy},{"left",x*25.4/dx},{"top",y*25.4/dy},{"printableWidth",pw*25.4/dx},{"printableHeight",ph*25.4/dy},{"dpiX",dx},{"dpiY",dy}};
                return result;
            } finally {if(context!=IntPtr.Zero)Native.DeleteDC(context);Native.Free(mode,bytes.Length);}
        }
        private string CapabilityName(IntPtr mode,object id,ushort idsCapability,ushort namesCapability,int nameLength) {
            if(id==null)return null;int count=Native.DeviceCapabilitiesW(Name,null,idsCapability,IntPtr.Zero,mode);
            if(count<=0||count>1024)return null;IntPtr ids=Native.Allocate(count*2),names=Native.Allocate(count*nameLength*2);
            try {int a=Native.DeviceCapabilitiesW(Name,null,idsCapability,ids,mode),b=Native.DeviceCapabilitiesW(Name,null,namesCapability,names,mode);
                if(a!=count||b!=count)return null;
                for(int n=0;n<count;n++)if((ushort)Marshal.ReadInt16(ids,n*2)==Convert.ToInt32(id))return Marshal.PtrToStringUni(IntPtr.Add(names,n*nameLength*2),nameLength).TrimEnd('\0');return null;
            } finally {Native.Free(ids,count*2);Native.Free(names,count*nameLength*2);}
        }
        private object Paper(IntPtr mode,object id,object orientation,byte[] bytes) {
            double w=0,h=0;uint fields=BitConverter.ToUInt32(bytes,72);
            if((fields&12)==12){w=BitConverter.ToInt16(bytes,82)/10.0;h=BitConverter.ToInt16(bytes,80)/10.0;}
            if((w<=0||h<=0)&&id!=null){int count=Native.DeviceCapabilitiesW(Name,null,2,IntPtr.Zero,mode);
                if(count>0&&count<=1024){IntPtr ids=Native.Allocate(count*2),sizes=Native.Allocate(count*8);
                    try {if(Native.DeviceCapabilitiesW(Name,null,2,ids,mode)==count&&Native.DeviceCapabilitiesW(Name,null,3,sizes,mode)==count)
                        for(int n=0;n<count;n++)if((ushort)Marshal.ReadInt16(ids,n*2)==Convert.ToInt32(id)){w=Marshal.ReadInt32(sizes,n*8)/10.0;h=Marshal.ReadInt32(sizes,n*8+4)/10.0;break;}}
                    finally {Native.Free(ids,count*2);Native.Free(sizes,count*8);}
                }
            }
            if(w<=0||h<=0)return null;if(orientation!=null&&Convert.ToInt32(orientation)==2){double swap=w;w=h;h=swap;}
            return new Dictionary<string,object>{{"name",CapabilityName(mode,id,2,16,64)},{"width",w},{"height",h}};
        }
        public void Dispose() {if(holdValidation!=null)holdValidation.Dispose();if(Mode!=null)Mode.Dispose();if(printer!=IntPtr.Zero){Native.ClosePrinter(printer);printer=IntPtr.Zero;}}
    }
}
