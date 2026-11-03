param([Parameter(Mandatory=$true)][string]$PdfPath)
$ErrorActionPreference='Stop'
if (-not [IO.File]::Exists($PdfPath) -or [IO.Path]::GetExtension($PdfPath) -ne '.pdf') { throw 'PDF de teste inválido.' }
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class PrintShopsOpen {
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] public struct Info {
  public int cbSize;public uint mask;public IntPtr hwnd;
  [MarshalAs(UnmanagedType.LPWStr)] public string verb;
  [MarshalAs(UnmanagedType.LPWStr)] public string file;
  [MarshalAs(UnmanagedType.LPWStr)] public string args;
  [MarshalAs(UnmanagedType.LPWStr)] public string directory;
  public int show;public IntPtr instance;public IntPtr idlist;
  [MarshalAs(UnmanagedType.LPWStr)] public string className;
  public IntPtr key;public uint hotkey;public IntPtr icon;public IntPtr process;
 }
 [DllImport("shell32.dll",CharSet=CharSet.Unicode,SetLastError=true)] public static extern bool ShellExecuteEx(ref Info info);
 public static void Open(string file){var info=new Info();info.cbSize=Marshal.SizeOf(info);info.mask=1;info.verb="open";info.file=file;info.className="br.com.4printshops.opentest.PDF";info.show=1;if(!ShellExecuteEx(ref info))throw new Exception("ShellExecuteEx: "+Marshal.GetLastWin32Error());}
}
'@
# Usa o ProgID de teste explicitamente; não altera nem invoca o leitor padrão.
[PrintShopsOpen]::Open($PdfPath)
