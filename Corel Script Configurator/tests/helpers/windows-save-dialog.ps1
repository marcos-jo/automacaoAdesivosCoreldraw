param([int]$DialogProcess, [string]$Destination, [ValidateSet('save','open','cancel-save','continue','discard','updates')] [string]$Action='save')
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class PrintShopsDialog {
 public delegate bool EnumProc(IntPtr h, IntPtr l);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr w,EnumProc p,IntPtr l);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr w,StringBuilder s,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr w,StringBuilder s,int n);
 [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr w);
 [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr w);
 [DllImport("user32.dll")] static extern bool PostMessage(IntPtr w,int m,IntPtr p,IntPtr l);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr w,int m,IntPtr p,string s);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr w,int m,IntPtr p,StringBuilder s);
 [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr w,int m,IntPtr p,IntPtr l);
 public static IntPtr Find(int pid,string title){IntPtr found=IntPtr.Zero;EnumWindows((h,l)=>{uint p;GetWindowThreadProcessId(h,out p);var s=new StringBuilder(512);GetWindowText(h,s,512);if(p==pid&&s.ToString()==title){found=h;return false;}return true;},IntPtr.Zero);return found;}
 public static string Titles(int pid){var result=new StringBuilder();EnumWindows((h,l)=>{uint p;GetWindowThreadProcessId(h,out p);if(p==pid){var s=new StringBuilder(512);GetWindowText(h,s,512);result.Append("["+s+"] ");}return true;},IntPtr.Zero);return result.ToString();}
 public static string Describe(IntPtr dialog){var result=new StringBuilder();EnumChildWindows(dialog,(h,l)=>{var c=new StringBuilder(128);GetClassName(h,c,128);if(c.ToString()=="Edit"||c.ToString()=="Button"||c.ToString()=="ComboBox"){var s=new StringBuilder(256);GetWindowText(h,s,256);result.Append(GetDlgCtrlID(h)+" "+c+" "+s+"; ");}return true;},IntPtr.Zero);return result.ToString();}
 public static bool Click(IntPtr dialog,string name){IntPtr button=IntPtr.Zero;EnumChildWindows(dialog,(h,l)=>{var c=new StringBuilder(128);GetClassName(h,c,128);var s=new StringBuilder(256);GetWindowText(h,s,256);if(c.ToString()=="Button"&&s.ToString().Replace("&","")==name)button=h;return true;},IntPtr.Zero);return button!=IntPtr.Zero&&PostMessage(button,245,IntPtr.Zero,IntPtr.Zero);}
 public static bool Save(IntPtr dialog,string destination){IntPtr edit=IntPtr.Zero,button=IntPtr.Zero;EnumChildWindows(dialog,(h,l)=>{int id=GetDlgCtrlID(h);var c=new StringBuilder(128);GetClassName(h,c,128);if((id==1001||id==1148)&&c.ToString()=="Edit")edit=h;if(id==1&&c.ToString()=="Button")button=h;return true;},IntPtr.Zero);if(edit==IntPtr.Zero||button==IntPtr.Zero)return false;var caption=new StringBuilder(128);GetWindowText(button,caption,128);string label=caption.ToString().Replace("&","");if(label!="Salvar"&&label!="Save"&&label!="Abrir"&&label!="Open")return false;SetFocus(edit);SendMessage(edit,12,IntPtr.Zero,destination);var actual=new StringBuilder(2048);SendMessage(edit,13,new IntPtr(2048),actual);if(actual.ToString()!=destination)return false;SendMessage(dialog,40,button,new IntPtr(1));SendMessage(button,245,IntPtr.Zero,IntPtr.Zero);return true;}
}
'@
$dialogRoot=[System.Windows.Automation.AutomationElement]::RootElement
$dialogEnd=(Get-Date).AddSeconds(20)
$dialogWindow=$null
$dialogTitle=switch($Action){'save'{'Salvar cópia PDF'}'cancel-save'{'Salvar cópia PDF'}'open'{'Abrir PDFs'}'updates'{'Atualizações do 4PrintShops'}default{'Alterações não exportadas'}}
while((Get-Date) -lt $dialogEnd -and -not $dialogWindow){
    $dialogHandle=[PrintShopsDialog]::Find($DialogProcess,$dialogTitle)
    if($dialogHandle -eq [IntPtr]::Zero -and $Action -eq 'open'){$dialogHandle=[PrintShopsDialog]::Find($DialogProcess,'Open')}
    if($dialogHandle -ne [IntPtr]::Zero){
        if($Action -in @('save','open') -and [PrintShopsDialog]::Save($dialogHandle,$Destination)){Write-Output 'Arquivo nativo confirmado por HWND/PID/título, campo Edit e destino verificado.';return}
        if($Action -notin @('save','open')){$dialogChoice=switch($Action){'continue'{'Continuar editando'}'cancel-save'{'Cancelar'}'updates'{'Continuar'}default{'Descartar e sair'}};if([PrintShopsDialog]::Click($dialogHandle,$dialogChoice)){return}}
        $dialogWindow=[System.Windows.Automation.AutomationElement]::FromHandle($dialogHandle)
        break
    }
    $dialogCondition=New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty,$DialogProcess)),
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty,$dialogTitle)))
    $dialogWindow=$dialogRoot.FindFirst([System.Windows.Automation.TreeScope]::Descendants,$dialogCondition)
    if(-not $dialogWindow){Start-Sleep -Milliseconds 100}
}
if(-not $dialogWindow){
    $dialogWindows=$dialogRoot.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)
    $dialogSeen=@();foreach($dialogSeenWindow in $dialogWindows){if($dialogSeenWindow.Current.ProcessId -eq $DialogProcess -or $dialogSeenWindow.Current.Name -match 'Salvar|4PrintShops'){$dialogSeen+=($dialogSeenWindow.Current.Name+' [PID '+$dialogSeenWindow.Current.ProcessId+']')}}
    throw ('Diálogo do processo de teste não encontrado. Win32: '+[PrintShopsDialog]::Titles($DialogProcess)+' Janelas relacionadas: '+($dialogSeen -join '; '))
}
if($Action -notin @('save','open')){
    $dialogButtonName=switch($Action){'continue'{'Continuar editando'}'cancel-save'{'Cancelar'}'updates'{'Continuar'}default{'Descartar e sair'}}
    $dialogButton=$dialogWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty,$dialogButtonName)))
    if(-not $dialogButton){throw 'Botão de confirmação não reconhecido.'}
    $dialogButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
    return
}
$dialogEdits=$dialogWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::Edit)))
$dialogName=$null;foreach($dialogEdit in $dialogEdits){if($dialogEdit.Current.AutomationId -in @('1001','1148')){$dialogName=$dialogEdit}}
$dialogButtons=$dialogWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::Button)))
$dialogSave=$null;foreach($dialogCandidate in $dialogButtons){if($dialogCandidate.Current.Name.Replace('&','') -in @('Salvar','Save','Abrir','Open')){$dialogSave=$dialogCandidate}}
if(-not $dialogName -or -not $dialogSave){
    $dialogControls=@();foreach($dialogEdit in $dialogEdits){$dialogControls+=('Edit '+$dialogEdit.Current.AutomationId+' '+$dialogEdit.Current.Name)};foreach($dialogCandidate in $dialogButtons){$dialogControls+=('Button '+$dialogCandidate.Current.AutomationId+' '+$dialogCandidate.Current.Name)}
    throw ('Controles nativos não reconhecidos; nenhuma tecla enviada. Win32: '+[PrintShopsDialog]::Describe($dialogHandle))
}
$dialogName.SetFocus()
$dialogName.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Destination)
$dialogSave.SetFocus()
$dialogSave.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
Write-Output 'Diálogo nativo confirmado exclusivamente na janela do processo de teste.'
