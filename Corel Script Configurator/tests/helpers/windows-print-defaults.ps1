param([Parameter(Mandatory=$true)][string]$PrinterName)
# Guarda bytes somente na memória do teste e devolve exclusivamente igualdade.
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class PrintConfigDefaultsTest {
 [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool OpenPrinterW(string n,out IntPtr h,IntPtr d);
 [DllImport("winspool.drv")] static extern bool ClosePrinter(IntPtr h);
 [DllImport("winspool.drv",CharSet=CharSet.Unicode)] static extern bool GetPrinterW(IntPtr h,uint l,IntPtr p,uint s,out uint n);
 public static byte[] Read(string name,uint level){IntPtr printer,buffer=IntPtr.Zero;uint needed;if(!OpenPrinterW(name,out printer,IntPtr.Zero))throw new Exception("Consulta de preferências indisponível.");try{GetPrinterW(printer,level,IntPtr.Zero,0,out needed);if(needed==0||needed>4194304)throw new Exception("Tamanho inválido.");buffer=Marshal.AllocHGlobal((int)needed);if(!GetPrinterW(printer,level,buffer,needed,out needed))throw new Exception("Consulta incompleta.");IntPtr mode=Marshal.ReadIntPtr(buffer);if(mode==IntPtr.Zero)return new byte[0];long offset=mode.ToInt64()-buffer.ToInt64();if(offset<0||offset+72>needed)throw new Exception("Ponteiro inválido.");int size=(ushort)Marshal.ReadInt16(mode,68)+(ushort)Marshal.ReadInt16(mode,70);if(size<72||offset+size>needed)throw new Exception("DEVMODE incompleto.");var data=new byte[size];Marshal.Copy(mode,data,0,size);return data;}finally{if(buffer!=IntPtr.Zero)Marshal.FreeHGlobal(buffer);ClosePrinter(printer);}}
 public static bool Same(byte[] a,byte[] b){a=a??new byte[0];b=b??new byte[0];try{if(a.Length!=b.Length)return false;for(int n=0;n<a.Length;n++)if(a[n]!=b[n])return false;return true;}finally{Array.Clear(b,0,b.Length);}}
}
'@
try {
    $globalBefore=[PrintConfigDefaultsTest]::Read($PrinterName,8)
    $userBefore=[PrintConfigDefaultsTest]::Read($PrinterName,9)
    Write-Output 'READY'
    while($null -ne ($watchCommand=[Console]::ReadLine())){
        if($watchCommand -eq 'quit'){break}
        if($watchCommand -ne 'check'){throw 'Comando inválido.'}
        $globalSame=[PrintConfigDefaultsTest]::Same($globalBefore,[PrintConfigDefaultsTest]::Read($PrinterName,8))
        $userSame=[PrintConfigDefaultsTest]::Same($userBefore,[PrintConfigDefaultsTest]::Read($PrinterName,9))
        if($globalSame -and $userSame){Write-Output 'UNCHANGED'}else{Write-Output 'CHANGED'}
    }
} finally {
    if($globalBefore){[Array]::Clear($globalBefore,0,$globalBefore.Length)}
    if($userBefore){[Array]::Clear($userBefore,0,$userBefore.Length)}
}
