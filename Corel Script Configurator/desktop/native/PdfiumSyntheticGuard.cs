using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace FourPrintShops.VirtualPrintLab {
    internal static partial class PdfiumVirtualLab {
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFPage_GetRotation(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFPage_CountObjects(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern IntPtr FPDFPage_GetObject(IntPtr page,int index);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFPageObj_GetType(IntPtr obj);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFPage_GetAnnotCount(IntPtr page);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern IntPtr FPDFBitmap_Create(int width,int height,int alpha);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDFBitmap_FillRect(IntPtr bitmap,int left,int top,int width,int height,uint color);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDF_RenderPageBitmap(IntPtr bitmap,IntPtr page,int x,int y,int width,int height,int rotation,int flags);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern IntPtr FPDFBitmap_GetBuffer(IntPtr bitmap);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern int FPDFBitmap_GetStride(IntPtr bitmap);
        [DllImport("pdfium.dll",CallingConvention=CallingConvention.Cdecl)] private static extern void FPDFBitmap_Destroy(IntPtr bitmap);

        // Bitmap usado somente para reconhecer o sintético. A impressão continua no PDF/HDC vetorial.
        internal static void RequireSynthetic(string candidate,string trusted) {
            if(!SetDefaultDllDirectories(0x200|0x800))throw new InvalidOperationException();
            FPDF_InitLibrary();
            try {
                if(!SyntheticPixels(candidate).SequenceEqual(SyntheticPixels(trusted)))throw new InvalidOperationException("O documento atual não corresponde ao sintético de teste.");
            } finally {FPDF_DestroyLibrary();}
        }
        private static byte[] SyntheticPixels(string file) {
            byte[] bytes=File.ReadAllBytes(file);
            var pinned=GCHandle.Alloc(bytes,GCHandleType.Pinned);
            IntPtr document=IntPtr.Zero,page=IntPtr.Zero,bitmap=IntPtr.Zero;
            try {
                document=FPDF_LoadMemDocument64(pinned.AddrOfPinnedObject(),new UIntPtr((uint)bytes.Length),null);
                if(document==IntPtr.Zero||FPDF_GetPageCount(document)!=1)throw new InvalidOperationException("O ensaio aceita somente o sintético de uma página.");
                page=FPDF_LoadPage(document,0);
                if(page==IntPtr.Zero||FPDFPage_GetAnnotCount(page)!=0||FPDFPage_CountObjects(page)!=3||
                    Math.Abs(FPDF_GetPageWidth(page)-595.2756)>0.1||Math.Abs(FPDF_GetPageHeight(page)-841.8898)>0.1)throw new InvalidOperationException("Use o sintético A4 sem anotações ou redimensionamento.");
                int[] types={1,1,2};
                for(int n=0;n<3;n++)if(FPDFPageObj_GetType(FPDFPage_GetObject(page,n))!=types[n])throw new InvalidOperationException("Conteúdo vetorial sintético incompatível.");
                int rotation=FPDFPage_GetRotation(page);
                if(rotation!=0&&rotation!=2)throw new InvalidOperationException("O ensaio aceita A4 retrato, incluindo giro de 180 graus.");
                bitmap=FPDFBitmap_Create(596,842,0);
                if(bitmap==IntPtr.Zero)throw new InvalidOperationException();
                FPDFBitmap_FillRect(bitmap,0,0,596,842,0xffffffff);
                FPDF_RenderPageBitmap(bitmap,page,0,0,596,842,rotation==2?2:0,0);
                byte[] pixels=new byte[checked(FPDFBitmap_GetStride(bitmap)*842)];
                Marshal.Copy(FPDFBitmap_GetBuffer(bitmap),pixels,0,pixels.Length);
                return pixels;
            } finally {
                if(bitmap!=IntPtr.Zero)FPDFBitmap_Destroy(bitmap);
                if(page!=IntPtr.Zero)FPDF_ClosePage(page);
                if(document!=IntPtr.Zero)FPDF_CloseDocument(document);
                pinned.Free();Array.Clear(bytes,0,bytes.Length);
            }
        }
    }
}
