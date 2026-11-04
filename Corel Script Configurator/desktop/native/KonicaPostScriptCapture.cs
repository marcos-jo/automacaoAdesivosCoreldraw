using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using FourPrintShops.PrintConfig;

namespace FourPrintShops.VirtualPrintLab {
    internal static partial class PdfiumVirtualLab {
        internal const string CaptureQueue = "4PrintShops Konica PS Capture Lab";

        private static void CheckCaptureQueue() {
            CaptureEnvironmentPolicy.RequireManualDefault();
            IntPtr handle;
            if (!OpenPrinterW(CaptureQueue, out handle, IntPtr.Zero)) throw new InvalidOperationException("Fila FILE: de laboratório ausente.");
            try {
                uint needed;
                GetPrinterW(handle, 2, IntPtr.Zero, 0, out needed);
                IntPtr buffer = Native.Allocate(checked((int)needed));
                int capacity = (int)needed;
                try {
                    if (!GetPrinterW(handle, 2, buffer, needed, out needed)) throw new InvalidOperationException();
                    var info = (Native.Printer2)Marshal.PtrToStructure(buffer, typeof(Native.Printer2));
                    if (Native.Text(info.Name) != CaptureQueue || Native.Text(info.Driver) != KonicaDriverIdentity.Driver ||
                        Native.Text(info.Port) != "FILE:" || !String.IsNullOrEmpty(Native.Text(info.Share)) ||
                        !String.IsNullOrEmpty(Native.Text(info.Server)) || (info.Attributes & 8) != 0) {
                        throw new InvalidOperationException("Captura recusada. A fila deve ser local, não compartilhada e exclusivamente FILE:.");
                    }
                } finally { Native.Free(buffer, capacity); }
            } finally { ClosePrinter(handle); }
        }

        internal static Dictionary<string, object> Capture(PrinterSession session, string input, string output, int scale, string allowedRoot = null, string jobName = null) {
            if (scale != 100 && scale != 120) throw new InvalidOperationException("Escala de laboratório inválida.");
            string root = Path.GetFullPath(allowedRoot ?? Path.Combine(Environment.CurrentDirectory, ".desktop-test", "hold-native")) + Path.DirectorySeparatorChar;
            output = Path.GetFullPath(output);
            if (!output.StartsWith(root, StringComparison.OrdinalIgnoreCase) || Path.GetExtension(output) != ".ps" || File.Exists(output)) {
                throw new InvalidOperationException("A captura exige arquivo PS novo na pasta de laboratório.");
            }
            CheckCaptureQueue();
            var summary = session.VerifyHold();
            byte[] bytes = File.ReadAllBytes(input), mode = session.Mode.Copy();
            IntPtr dc = IntPtr.Zero, memory = IntPtr.Zero, document = IntPtr.Zero, page = IntPtr.Zero;
            GCHandle pinned = default(GCHandle);
            bool initialized = false, started = false;
            try {
                if (bytes.Length < 8 || bytes.Length > 1024 * 1024) throw new InvalidOperationException("Use somente o sintético de laboratório.");
                if (!SetDefaultDllDirectories(0x200 | 0x800)) throw new InvalidOperationException();
                FPDF_InitLibrary(); initialized = true;
                pinned = GCHandle.Alloc(bytes, GCHandleType.Pinned);
                document = FPDF_LoadMemDocument64(pinned.AddrOfPinnedObject(), new UIntPtr((uint)bytes.Length), null);
                if (document == IntPtr.Zero || FPDF_GetPageCount(document) != 1) throw new InvalidOperationException("A captura exige uma página sintética.");
                page = FPDF_LoadPage(document, 0);
                if (page == IntPtr.Zero) throw new InvalidOperationException();
                double sourceWidth = FPDF_GetPageWidth(page), sourceHeight = FPDF_GetPageHeight(page);
                if (Math.Abs(sourceWidth - 595.2756) > 0.1 || Math.Abs(sourceHeight - 841.8898) > 0.1) throw new InvalidOperationException("Origem A4 retrato exigida.");
                memory = Native.Allocate(mode.Length); Marshal.Copy(mode, 0, memory, mode.Length);
                dc = CreateDCW(null, CaptureQueue, null, memory);
                if (dc == IntPtr.Zero) throw new InvalidOperationException("O driver não criou o contexto FILE:.");
                byte[] after = Native.ReadMode(memory, mode.Length);
                try { if (!mode.SequenceEqual(after)) throw new InvalidOperationException("O driver alterou o buffer confirmado."); }
                finally { Array.Clear(after, 0, after.Length); }
                int dpiX = GetDeviceCaps(dc, 88), dpiY = GetDeviceCaps(dc, 90);
                if (dpiX <= 0 || dpiY <= 0) throw new InvalidOperationException("Métricas inválidas.");
                int paperWidth = GetDeviceCaps(dc, 110), paperHeight = GetDeviceCaps(dc, 111);
                int width = (int)Math.Round(sourceWidth * dpiX / 72.0 * scale / 100.0);
                int height = (int)Math.Round(sourceHeight * dpiY / 72.0 * scale / 100.0);
                int x = (paperWidth - width) / 2 - GetDeviceCaps(dc, 112);
                int y = (paperHeight - height) / 2 - GetDeviceCaps(dc, 113);
                // As métricas físicas do PScript arredondam A4 em cerca de 0,09 mm nesta estação.
                int toleranceX = (int)Math.Ceiling(dpiX * 0.1 / 25.4), toleranceY = (int)Math.Ceiling(dpiY * 0.1 / 25.4);
                if (width > paperWidth + toleranceX || height > paperHeight + toleranceY) throw new InvalidOperationException("O conteúdo ultrapassa o papel confirmado.");
                if (FPDF_SetPrintMode(0) == 0) throw new InvalidOperationException("Modo EMF indisponível.");
                CheckCaptureQueue();
                var info = new DocumentInfo { Size = Marshal.SizeOf(typeof(DocumentInfo)), Name = jobName ?? "4PS PS CAPTURE " + Guid.NewGuid().ToString("N").Substring(0, 8), Output = output };
                if (StartDocW(dc, ref info) <= 0) throw new InvalidOperationException("A captura FILE: não iniciou.");
                started = true;
                if (StartPage(dc) <= 0 || FPDF_RenderPage(dc, page, x, y, width, height, 0, 0x801) == 0 ||
                    EndPage(dc) <= 0 || EndDoc(dc) <= 0) throw new InvalidOperationException("A captura FILE: não foi concluída.");
                started = false;
                summary["captureQueue"] = CaptureQueue; summary["output"] = output; summary["scale"] = scale;
                summary["contextWidthMM"] = paperWidth * 25.4 / dpiX; summary["contextHeightMM"] = paperHeight * 25.4 / dpiY;
                summary["contentWidthMM"] = width * 25.4 / dpiX; summary["contentHeightMM"] = height * 25.4 / dpiY;
                return summary;
            } finally {
                if (started && dc != IntPtr.Zero) AbortDoc(dc);
                if (dc != IntPtr.Zero) DeleteDC(dc);
                if (memory != IntPtr.Zero) Native.Free(memory, mode.Length);
                if (page != IntPtr.Zero) FPDF_ClosePage(page);
                if (document != IntPtr.Zero) FPDF_CloseDocument(document);
                if (pinned.IsAllocated) pinned.Free();
                if (initialized) FPDF_DestroyLibrary();
                Array.Clear(bytes, 0, bytes.Length); Array.Clear(mode, 0, mode.Length);
            }
        }
    }

    internal static class KonicaCaptureHost {
        [STAThread] private static int Main(string[] args) {
            Console.OutputEncoding = new UTF8Encoding(false);
            try {
                if (args.Length != 3) throw new InvalidOperationException("Use sintético, saída PS e escala 100/120.");
                using (var session = new PrinterSession(KonicaDriverIdentity.Queue)) {
                    Console.WriteLine("PROPERTIES_READY");
                    var state = session.Properties(IntPtr.Zero, session.Mode.Revision);
                    if (state.ContainsKey("canceled") && (bool)state["canceled"]) throw new InvalidOperationException("Captura cancelada no diálogo.");
                    var result = PdfiumVirtualLab.Capture(session, args[0], args[1], Int32.Parse(args[2]));
                    Console.WriteLine(new JavaScriptSerializer().Serialize(result));
                }
                return 0;
            } catch (Exception error) {
                Console.Error.WriteLine(error is InvalidOperationException ? error.Message : "Falha na captura isolada. Nenhum destino físico é permitido.");
                return 1;
            }
        }
    }
}
