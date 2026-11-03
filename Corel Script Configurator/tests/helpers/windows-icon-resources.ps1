param([Parameter(Mandatory=$true)][string]$Installation,[Parameter(Mandatory=$true)][string]$Installer,[Parameter(Mandatory=$true)][string]$Output,[Parameter(Mandatory=$true)][string]$ProjectRoot)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
# Somente leitura de recursos/registro/atalhos; sem automação da interface Windows.
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;using System.Drawing;using System.Drawing.Imaging;using System.Runtime.InteropServices;using System.Text;
public static class IconResources {
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern uint PrivateExtractIcons(string file,int index,int width,int height,IntPtr[] icons,uint[] ids,uint count,uint flags);
 [DllImport("user32.dll")]static extern bool DestroyIcon(IntPtr handle);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode)]static extern uint GetLongPathName(string path,StringBuilder output,uint size);
 public static string LongPath(string path){var result=new StringBuilder(4096);if(GetLongPathName(path,result,4096)==0)throw new Exception("Caminho nao encontrado: "+path);return result.ToString();}
 static Bitmap Extract(string file,int size){var icons=new IntPtr[1];var ids=new uint[1];if(PrivateExtractIcons(file,0,size,size,icons,ids,1,0)!=1||icons[0]==IntPtr.Zero)throw new Exception("Icone ausente em "+file);try{using(var icon=(Icon)Icon.FromHandle(icons[0]).Clone())return icon.ToBitmap();}finally{DestroyIcon(icons[0]);}}
 public static void Compare(string file,string expected,int size,string output){using(var actual=Extract(file,size))using(var reference=Extract(expected,size)){if(actual.Width!=size||actual.Height!=size)throw new Exception("Dimensao incorreta");for(int y=0;y<size;y++)for(int x=0;x<size;x++)if(actual.GetPixel(x,y)!=reference.GetPixel(x,y))throw new Exception("Recurso visual diferente: "+file+" em "+size);actual.Save(output,ImageFormat.Png);}}
}
'@
[IO.Directory]::CreateDirectory($Output)|Out-Null
$appIcon=Join-Path $ProjectRoot 'build/icons/app.ico'
$pdfIcon=Join-Path $ProjectRoot 'build/icons/pdf.ico'
$exe=Join-Path $Installation '4PrintShops Icon Test.exe'
foreach($entry in @(@($exe,$appIcon,'executavel'),@($Installer,$appIcon,'instalador'),@((Join-Path $Installation 'Uninstall 4PrintShops Icon Test.exe'),$appIcon,'desinstalador'),@((Join-Path $Installation 'resources/icons/pdf.ico'),$pdfIcon,'pdf'))){
    foreach($size in @(16,32)){[IconResources]::Compare($entry[0],$entry[1],$size,(Join-Path $Output ($entry[2]+'-'+$size+'.png')))}
}
$registry=[Microsoft.Win32.Registry]::CurrentUser
$pdfKey=$registry.OpenSubKey('Software\Classes\br.com.4printshops.icontest.PDF\DefaultIcon')
try{if($pdfKey.GetValue('') -ne ('"'+(Join-Path $Installation 'resources\icons\pdf.ico')+'",0')){throw 'DefaultIcon PDF incorreto'}}finally{$pdfKey.Close()}
$command=$registry.OpenSubKey('Software\Classes\br.com.4printshops.icontest.PDF\shell\open\command')
try{if($command.GetValue('') -ne ('"'+$exe+'" --open-pdf "%1"')){throw 'Comando de abertura alterado'}}finally{$command.Close()}
$shell=New-Object -ComObject WScript.Shell
foreach($shortcut in @((Join-Path ([Environment]::GetFolderPath('Desktop')) '4PrintShops Icon Test.lnk'),(Join-Path ([Environment]::GetFolderPath('Programs')) '4PrintShops Icon Test.lnk'))){
    if(-not (Test-Path -LiteralPath $shortcut)){throw "Atalho nao encontrado: $shortcut"}
    $link=$shell.CreateShortcut($shortcut)
    if([IconResources]::LongPath($link.TargetPath) -ne [IconResources]::LongPath($exe)){throw "Destino de atalho incorreto: $shortcut"}
    if($link.IconLocation -and [IconResources]::LongPath(($link.IconLocation -replace ',0$','').Trim('"')) -ne [IconResources]::LongPath($exe)){throw "Icone de atalho incorreto: $shortcut"}
}
Write-Output 'PASSOU: icones reais de executavel/Setup/desinstalador/PDF em 16 e 32 iguais aos ICOs; ProgID PDF e atalhos corretos.'
