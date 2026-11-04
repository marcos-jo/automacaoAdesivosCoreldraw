using System;
using System.IO;
using System.Runtime.InteropServices;

namespace FourPrintShops.VirtualPrintLab {
    internal static partial class PdfiumVirtualLab {
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDF_CreateNewDocument();
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDFPage_New(IntPtr document, int index, double width, double height);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDFText_LoadFont(IntPtr document, byte[] data, uint length, int type, int cid);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDFPageObj_CreateTextObj(IntPtr document, IntPtr font, float size);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDFText_SetText(IntPtr text, [MarshalAs(UnmanagedType.LPWStr)] string value);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern void FPDFPageObj_Transform(IntPtr obj, double a, double b, double c, double d, double e, double f);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern void FPDFPage_InsertObject(IntPtr page, IntPtr obj);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDFPage_GenerateContent(IntPtr page);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern IntPtr FPDFPageObj_CreateNewRect(float x, float y, float width, float height);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDFPageObj_SetFillColor(IntPtr obj, uint r, uint g, uint b, uint a);
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDFPath_SetDrawMode(IntPtr path, int fill, int stroke);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate int WriteBlock(IntPtr writer, IntPtr data, uint size);
        [StructLayout(LayoutKind.Sequential)] private struct FileWriter { public int Version; public WriteBlock Write; }
        [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)] private static extern int FPDF_SaveAsCopy(IntPtr document, ref FileWriter writer, uint flags);

        private static void CreateFixture(string fontPath, string output) {
            if (File.Exists(output)) throw new InvalidOperationException("A saída sintética deve ser nova.");
            if (!SetDefaultDllDirectories(0x200 | 0x800)) throw new InvalidOperationException();
            FPDF_InitLibrary();
            IntPtr document = IntPtr.Zero, page = IntPtr.Zero;
            byte[] fontBytes = File.ReadAllBytes(fontPath);
            try {
                document = FPDF_CreateNewDocument();
                page = FPDFPage_New(document, 0, 595.2756, 841.8898);
                IntPtr font = FPDFText_LoadFont(document, fontBytes, (uint)fontBytes.Length, 2, 1);
                if (document == IntPtr.Zero || page == IntPtr.Zero || font == IntPtr.Zero) throw new InvalidOperationException("Fonte sintética indisponível.");
                AddText(document, page, font, "4PrintShops HOLD TEST - SYNTHETIC PAGE", 18, 36, 770);
                AddText(document, page, font, "PDFium native / embedded TrueType / vectors", 12, 36, 735);
                IntPtr rectangle = FPDFPageObj_CreateNewRect(36, 620, 150, 50);
                if (rectangle == IntPtr.Zero || FPDFPageObj_SetFillColor(rectangle, 51, 102, 204, 255) == 0 || FPDFPath_SetDrawMode(rectangle, 2, 0) == 0) throw new InvalidOperationException();
                FPDFPage_InsertObject(page, rectangle);
                if (FPDFPage_GenerateContent(page) == 0) throw new InvalidOperationException();
                using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    WriteBlock write = delegate(IntPtr writer, IntPtr data, uint size) {
                        try {
                            var buffer = new byte[size];
                            Marshal.Copy(data, buffer, 0, buffer.Length);
                            stream.Write(buffer, 0, buffer.Length);
                            return 1;
                        } catch { return 0; }
                    };
                    var callbacks = new FileWriter { Version = 1, Write = write };
                    if (FPDF_SaveAsCopy(document, ref callbacks, 0) == 0) throw new InvalidOperationException("A página sintética não foi gravada.");
                    GC.KeepAlive(write);
                }
            } finally {
                if (page != IntPtr.Zero) FPDF_ClosePage(page);
                if (document != IntPtr.Zero) FPDF_CloseDocument(document);
                FPDF_DestroyLibrary();
                Array.Clear(fontBytes, 0, fontBytes.Length);
            }
        }

        private static void AddText(IntPtr document, IntPtr page, IntPtr font, string text, float size, double x, double y) {
            IntPtr obj = FPDFPageObj_CreateTextObj(document, font, size);
            if (obj == IntPtr.Zero || FPDFText_SetText(obj, text) == 0) throw new InvalidOperationException("Texto sintético indisponível.");
            FPDFPageObj_Transform(obj, 1, 0, 0, 1, x, y);
            FPDFPage_InsertObject(page, obj);
        }
    }
}
