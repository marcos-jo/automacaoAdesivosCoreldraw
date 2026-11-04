param([Parameter(Mandatory=$true)][int]$DriverProcess,[ValidateSet('inspect','cancel','accept','a3','a3-sizes','landscape','workflow','hold','print','holdprint')][string]$Action='inspect')
# Somente o diálogo de propriedades do auxiliar de teste; nunca abre impressoras ou trabalhos.
Add-Type -TypeDefinition @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class PrintConfigDialogTest {
 public delegate bool EnumProc(IntPtr h,IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f,IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr w,EnumProc f,IntPtr p);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr w,StringBuilder s,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr w,StringBuilder s,int n);
 [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr w);
 [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr w);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr w);
 [DllImport("user32.dll")] static extern bool PostMessage(IntPtr w,uint m,IntPtr a,IntPtr b);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr w,uint m,IntPtr a,IntPtr b);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr w,uint m,IntPtr a,StringBuilder b);
 static string Class(IntPtr w){var s=new StringBuilder(80);GetClassName(w,s,80);return s.ToString();}
 static string Text(IntPtr w){var s=new StringBuilder(512);GetWindowText(w,s,512);return s.ToString();}
 public static IntPtr Find(int pid){IntPtr found=IntPtr.Zero;EnumWindows((w,p)=>{uint current;GetWindowThreadProcessId(w,out current);if(current==pid&&IsWindowVisible(w)&&Class(w)=="#32770"){found=w;return false;}return true;},IntPtr.Zero);return found;}
 public static bool Click(IntPtr dialog,int id){IntPtr button=IntPtr.Zero;EnumChildWindows(dialog,(w,p)=>{if(GetDlgCtrlID(w)==id&&Class(w)=="Button"&&IsWindowVisible(w))button=w;return true;},IntPtr.Zero);return button!=IntPtr.Zero&&PostMessage(button,0x00F5,IntPtr.Zero,IntPtr.Zero);}
 public static string Inspect(IntPtr dialog){var result=new List<string>();EnumChildWindows(dialog,(w,p)=>{string c=Class(w);if(IsWindowVisible(w)&&(c=="Button"||c=="ComboBox"))result.Add(GetDlgCtrlID(w)+" "+c+" "+Text(w));return true;},IntPtr.Zero);return String.Join("; ",result.ToArray());}
 public static bool SelectA3(IntPtr dialog){IntPtr combo=IntPtr.Zero;int index=-1;EnumChildWindows(dialog,(w,p)=>{if(Class(w)!="ComboBox"||!IsWindowVisible(w))return true;int count=SendMessage(w,0x0146,IntPtr.Zero,IntPtr.Zero).ToInt32();if(count<1||count>512)return true;for(int n=0;n<count;n++){var text=new StringBuilder(512);SendMessage(w,0x0148,new IntPtr(n),text);if(text.ToString()=="A3"){combo=w;index=n;return false;}}return true;},IntPtr.Zero);if(combo==IntPtr.Zero)return false;SendMessage(combo,0x014E,new IntPtr(index),IntPtr.Zero);int id=GetDlgCtrlID(combo);PostMessage(dialog,0x0111,new IntPtr((1<<16)|(id&0xffff)),combo);return true;}
 public static bool Landscape(IntPtr dialog){IntPtr control=IntPtr.Zero;int index=-1;EnumChildWindows(dialog,(w,p)=>{if(!IsWindowVisible(w))return true;string c=Class(w);if(c=="Button"&&(Text(w).Replace("&","")=="Landscape"||Text(w).Replace("&","")=="Paisagem")){control=w;return false;}if(c=="ComboBox"){int count=SendMessage(w,0x0146,IntPtr.Zero,IntPtr.Zero).ToInt32();if(count<1||count>512)return true;for(int n=0;n<count;n++){var text=new StringBuilder(512);SendMessage(w,0x0148,new IntPtr(n),text);if(text.ToString()=="Paisagem"||text.ToString()=="Landscape"){control=w;index=n;return false;}}}return true;},IntPtr.Zero);if(control==IntPtr.Zero)return false;if(index<0)return PostMessage(control,0x00F5,IntPtr.Zero,IntPtr.Zero);SendMessage(control,0x014E,new IntPtr(index),IntPtr.Zero);int id=GetDlgCtrlID(control);return PostMessage(dialog,0x0111,new IntPtr((1<<16)|(id&0xffff)),control);}
 public static bool Workflow(IntPtr dialog){IntPtr button=IntPtr.Zero;EnumChildWindows(dialog,(w,p)=>{if(IsWindowVisible(w)&&Class(w)=="Button"&&Text(w).Replace("&","")=="Workflow")button=w;return true;},IntPtr.Zero);return button!=IntPtr.Zero&&PostMessage(button,0x00F5,IntPtr.Zero,IntPtr.Zero);}
 public static bool A3Sizes(IntPtr dialog){int changed=0;EnumChildWindows(dialog,(w,p)=>{if(!IsWindowVisible(w)||Class(w)!="ComboBox"||GetDlgCtrlID(w)==-31972)return true;int count=SendMessage(w,0x0146,IntPtr.Zero,IntPtr.Zero).ToInt32();if(count<2||count>512)return true;bool a4=false;int a3=-1;for(int n=0;n<count;n++){var value=new StringBuilder(512);SendMessage(w,0x0148,new IntPtr(n),value);if(value.ToString()=="A4")a4=true;if(value.ToString()=="A3")a3=n;}if(a4&&a3>=0){SendMessage(w,0x014E,new IntPtr(a3),IntPtr.Zero);int id=GetDlgCtrlID(w)&0xffff;SendMessage(GetParent(w),0x0111,new IntPtr((1<<16)|id),w);SendMessage(GetParent(w),0x0111,new IntPtr((9<<16)|id),w);changed++;}return true;},IntPtr.Zero);return changed>0;}
 public static bool OutputMethod(IntPtr dialog,string action){IntPtr selected=IntPtr.Zero;int selectedIndex=-1,matches=0;EnumChildWindows(dialog,(w,p)=>{if(!IsWindowVisible(w)||Class(w)!="ComboBox")return true;int count=SendMessage(w,0x0146,IntPtr.Zero,IntPtr.Zero).ToInt32();if(count<3||count>100)return true;bool hasHold=false,hasPrint=false;int target=-1;for(int n=0;n<count;n++){var text=new StringBuilder(512);SendMessage(w,0x0148,new IntPtr(n),text);string value=text.ToString().Trim();if(value=="Hold"||value=="Reter")hasHold=true;if(value=="Print"||value=="Imprimir")hasPrint=true;if(action=="hold"&&(value=="Hold"||value=="Reter")||action=="print"&&(value=="Print"||value=="Imprimir")||action=="holdprint"&&(value=="Print/Hold"||value=="Imprimir/Reter"))target=n;}if(hasHold&&hasPrint&&target>=0){matches++;selected=w;selectedIndex=target;}return true;},IntPtr.Zero);if(matches!=1)return false;SendMessage(selected,0x014E,new IntPtr(selectedIndex),IntPtr.Zero);int id=GetDlgCtrlID(selected);SendMessage(GetParent(selected),0x0111,new IntPtr((1<<16)|(id&0xffff)),selected);return PostMessage(GetParent(selected),0x0111,new IntPtr((9<<16)|(id&0xffff)),selected);}
}
'@
$dialogDeadline=(Get-Date).AddSeconds(20)
$driverDialog=[IntPtr]::Zero
while((Get-Date) -lt $dialogDeadline -and $driverDialog -eq [IntPtr]::Zero){$driverDialog=[PrintConfigDialogTest]::Find($DriverProcess);if($driverDialog -eq [IntPtr]::Zero){Start-Sleep -Milliseconds 100}}
if($driverDialog -eq [IntPtr]::Zero){throw 'Diálogo de propriedades do auxiliar não encontrado.'}
switch($Action){
 'inspect'{[PrintConfigDialogTest]::Inspect($driverDialog)}
 'cancel'{if(-not [PrintConfigDialogTest]::Click($driverDialog,2)){throw 'Cancelar não encontrado no diálogo.'};Write-Output 'Cancelamento solicitado ao diálogo nativo.'}
 'accept'{if(-not [PrintConfigDialogTest]::Click($driverDialog,1)){throw 'OK não encontrado no diálogo.'};Write-Output 'Confirmação solicitada ao diálogo nativo.'}
 'a3'{if(-not [PrintConfigDialogTest]::SelectA3($driverDialog)){throw 'Opção A3 não encontrada em um campo padrão do diálogo.'};Write-Output 'A3 selecionado no diálogo nativo.'}
 'a3-sizes'{if(-not [PrintConfigDialogTest]::A3Sizes($driverDialog)){throw 'Campos A3/A4 não identificados.'};Write-Output 'Tamanhos A3 selecionados somente no diálogo de laboratório.'}
 'landscape'{if(-not [PrintConfigDialogTest]::Landscape($driverDialog)){throw 'Orientação paisagem não encontrada no diálogo.'};Write-Output 'Paisagem selecionada somente no diálogo nativo.'}
 'workflow'{if(-not [PrintConfigDialogTest]::Workflow($driverDialog)){throw 'Guia Workflow não encontrada.'};Write-Output 'Guia Workflow selecionada.'}
 {$_ -in @('hold','print','holdprint')}{if(-not [PrintConfigDialogTest]::OutputMethod($driverDialog,$Action)){throw 'Campo Output Method não identificado inequivocamente.'};Write-Output 'Output Method alterado somente no diálogo de propriedades.'}
}
