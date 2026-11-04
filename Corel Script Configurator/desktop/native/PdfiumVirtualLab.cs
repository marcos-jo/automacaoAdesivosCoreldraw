using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace FourPrintShops.VirtualPrintLab {
    // Este executável aceita exclusivamente o driver virtual Microsoft com saída explícita para arquivo.
    internal static partial class PdfiumVirtualLab {
        private const string Queue = "Microsoft Print to PDF";
        private const string Driver = "Microsoft Print To PDF";

        [StructLayout(LayoutKind.Sequential)]
        private struct PrinterInfo2 {
            public IntPtr Server, Name, Share, Port, Driver, Comment, Location, Mode, Separator, Processor, Datatype, Parameters, Security;
            public uint Attributes, Priority, DefaultPriority, StartTime, UntilTime, Status, Jobs, AveragePPM;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DocumentInfo {
            public int Size;
            public string Name;
            public string Output;
            public string Datatype;
            public uint Type;
        }

        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetDefaultDllDirectories(uint flags);
        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool OpenPrinterW(string name, out IntPtr printer, IntPtr defaults);
        [DllImport("winspool.drv")] private static extern bool ClosePrinter(IntPtr printer);
        [DllImport("winspool.drv", CharSet = CharSet.Unicode)] private static extern bool GetPrinterW(IntPtr printer, uint level, IntPtr buffer, uint size, out uint needed);
        [DllImport("winspool.drv", CharSet = CharSet.Unicode)] private static extern int DocumentPropertiesW(IntPtr owner, IntPtr printer, string name, IntPtr output, IntPtr input, uint flags);
        [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateDCW(string driver, string name, string port, IntPtr mode);
        [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr context);
        [DllImport("gdi32.dll")] private static extern int GetDeviceCaps(IntPtr context, int capability);
        [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern int StartDocW(IntPtr context, ref DocumentInfo info);
        [DllImport("gdi32.dll")] private static extern int StartPage(IntPtr context);
        [DllImport("gdi32.dll")] private static extern int EndPage(IntPtr context);
        [DllImport("gdi32.dll")] private static extern int EndDoc(IntPtr context);
        [DllImport("gdi32.dll")] private static extern int AbortDoc(IntPtr context);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern void FPDF_InitLibrary();
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern void FPDF_DestroyLibrary();
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDF_LoadMemDocument64(IntPtr data, UIntPtr length, string password);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDF_GetPageCount(IntPtr document);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDF_LoadPage(IntPtr document, int index);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern double FPDF_GetPageWidth(IntPtr page);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern double FPDF_GetPageHeight(IntPtr page);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDF_RenderPage(IntPtr context, IntPtr page, int x, int y, int width, int height, int rotation, int flags);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDF_SetPrintMode(int mode);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern void FPDF_ClosePage(IntPtr page);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern void FPDF_CloseDocument(IntPtr document);

        private static IntPtr ReadVirtualConfiguration(out IntPtr printer) {
            if (!OpenPrinterW(Queue, out printer, IntPtr.Zero)) throw new InvalidOperationException("A impressora virtual Microsoft Print to PDF está indisponível.");
            uint needed;
            GetPrinterW(printer, 2, IntPtr.Zero, 0, out needed);
            if (needed == 0 || needed > 4 * 1024 * 1024) throw new InvalidOperationException("Identidade virtual inválida.");
            IntPtr metadata = Marshal.AllocHGlobal((int)needed);
            try {
                if (!GetPrinterW(printer, 2, metadata, needed, out needed)) throw new InvalidOperationException("Consulta virtual incompleta.");
                var info = (PrinterInfo2)Marshal.PtrToStructure(metadata, typeof(PrinterInfo2));
                if (Marshal.PtrToStringUni(info.Name) != Queue ||
                    !String.Equals(Marshal.PtrToStringUni(info.Driver), Driver, StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(Marshal.PtrToStringUni(info.Port), "PORTPROMPT:", StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException("Fila, driver ou porta não correspondem ao destino virtual permitido.");
                }
            } finally {
                Marshal.FreeHGlobal(metadata);
            }
            int capacity = DocumentPropertiesW(IntPtr.Zero, printer, Queue, IntPtr.Zero, IntPtr.Zero, 0);
            if (capacity < 220 || capacity > 4 * 1024 * 1024) throw new InvalidOperationException("DEVMODE virtual inválido.");
            IntPtr mode = Marshal.AllocHGlobal(capacity);
            try {
                for (int n = 0; n < capacity; n++) Marshal.WriteByte(mode, n, 0);
                if (DocumentPropertiesW(IntPtr.Zero, printer, Queue, mode, IntPtr.Zero, 2) != 1) throw new InvalidOperationException();
                int total = (ushort)Marshal.ReadInt16(mode, 68) + (ushort)Marshal.ReadInt16(mode, 70);
                if (total > capacity || total < 220) throw new InvalidOperationException();
                int fields = Marshal.ReadInt32(mode, 72);
                Marshal.WriteInt32(mode, 72, fields | 1 | 2 | 256);
                Marshal.WriteInt16(mode, 76, 1); // Retrato, somente para este trabalho virtual.
                Marshal.WriteInt16(mode, 78, 9); // A4.
                Marshal.WriteInt16(mode, 86, 1); // Uma cópia.
                if (DocumentPropertiesW(IntPtr.Zero, printer, Queue, mode, mode, 10) != 1) throw new InvalidOperationException();
                if (Marshal.ReadInt16(mode, 76) != 1 || Marshal.ReadInt16(mode, 78) != 9 || Marshal.ReadInt16(mode, 86) != 1) {
                    throw new InvalidOperationException("O driver virtual alterou papel, orientação ou cópias.");
                }
                return mode;
            } catch {
                Marshal.FreeHGlobal(mode);
                throw;
            }
        }

        private static void Render(string input, string output) {
            if (!Path.IsPathRooted(input) || !Path.IsPathRooted(output) ||
                !String.Equals(Path.GetExtension(output), ".pdf", StringComparison.OrdinalIgnoreCase) || File.Exists(output)) {
                throw new InvalidOperationException("O laboratório requer entrada local e saída PDF nova.");
            }
            byte[] bytes = File.ReadAllBytes(input);
            if (bytes.Length < 8 || bytes.Length > 1024 * 1024) throw new InvalidOperationException("Use o PDF sintético de uma página.");
            IntPtr printer = IntPtr.Zero, mode = IntPtr.Zero, context = IntPtr.Zero, document = IntPtr.Zero, page = IntPtr.Zero;
            GCHandle pinned = default(GCHandle);
            bool initialized = false, started = false;
            try {
                FourPrintShops.PrintConfig.CaptureEnvironmentPolicy.RequireManualDefault();
                mode = ReadVirtualConfiguration(out printer);
                if (!SetDefaultDllDirectories(0x200 | 0x800)) throw new InvalidOperationException("Busca segura de DLL indisponível.");
                FPDF_InitLibrary();
                initialized = true;
                pinned = GCHandle.Alloc(bytes, GCHandleType.Pinned);
                document = FPDF_LoadMemDocument64(pinned.AddrOfPinnedObject(), new UIntPtr((uint)bytes.Length), null);
                if (document == IntPtr.Zero || FPDF_GetPageCount(document) != 1) throw new InvalidOperationException("O laboratório aceita uma única página sintética.");
                page = FPDF_LoadPage(document, 0);
                if (page == IntPtr.Zero) throw new InvalidOperationException("PDFium não abriu a página.");
                context = CreateDCW(null, Queue, null, mode);
                if (context == IntPtr.Zero) throw new InvalidOperationException("Contexto virtual indisponível.");
                int dpiX = GetDeviceCaps(context, 88), dpiY = GetDeviceCaps(context, 90);
                int physicalWidth = GetDeviceCaps(context, 110), physicalHeight = GetDeviceCaps(context, 111);
                double paperWidth = physicalWidth * 25.4 / dpiX, paperHeight = physicalHeight * 25.4 / dpiY;
                if (dpiX <= 0 || dpiY <= 0 || Math.Abs(paperWidth - 210) > 1 || Math.Abs(paperHeight - 297) > 1) {
                    throw new InvalidOperationException("O contexto virtual não conservou o A4 confirmado.");
                }
                int width = (int)Math.Round(FPDF_GetPageWidth(page) * dpiX / 72.0);
                int height = (int)Math.Round(FPDF_GetPageHeight(page) * dpiY / 72.0);
                var info = new DocumentInfo { Size = Marshal.SizeOf(typeof(DocumentInfo)), Name = "4PrintShops VIRTUAL HOLD TEST", Output = output };
                int job = StartDocW(context, ref info);
                if (job <= 0) throw new InvalidOperationException("O trabalho virtual não iniciou.");
                started = true;
                if (StartPage(context) <= 0) throw new InvalidOperationException("A página virtual não iniciou.");
                if (FPDF_SetPrintMode(0) == 0) throw new InvalidOperationException("Modo EMF do PDFium indisponível.");
                int x = -GetDeviceCaps(context, 112), y = -GetDeviceCaps(context, 113);
                if (FPDF_RenderPage(context, page, x, y, width, height, 0, 0x800 | 1) == 0) throw new InvalidOperationException("PDFium falhou ao gerar a página virtual.");
                if (EndPage(context) <= 0 || EndDoc(context) <= 0) throw new InvalidOperationException("A saída virtual não foi concluída.");
                started = false;
                Console.WriteLine("Trabalho virtual concluído. Destino Microsoft Print to PDF. Job " + job + ".");
            } finally {
                if (started && context != IntPtr.Zero) AbortDoc(context);
                if (context != IntPtr.Zero) DeleteDC(context);
                if (page != IntPtr.Zero) FPDF_ClosePage(page);
                if (document != IntPtr.Zero) FPDF_CloseDocument(document);
                if (pinned.IsAllocated) pinned.Free();
                if (initialized) FPDF_DestroyLibrary();
                if (mode != IntPtr.Zero) Marshal.FreeHGlobal(mode);
                if (printer != IntPtr.Zero) ClosePrinter(printer);
                Array.Clear(bytes, 0, bytes.Length);
            }
        }

        [STAThread]
        private static int Main(string[] args) {
            try {
                if (args.Length != 3) throw new InvalidOperationException("Modo de laboratório inválido.");
                if (args[0] == "--create-synthetic") CreateFixture(args[1], args[2]);
                else if (args[0] == "--virtual-file") Render(args[1], args[2]);
                else throw new InvalidOperationException("Modo de laboratório inválido.");
                return 0;
            } catch (Exception error) {
                Console.Error.WriteLine(error is InvalidOperationException ? error.Message : "Falha no laboratório PDFium. Nenhum destino físico é permitido.");
                return 1;
            }
        }
    }
}
