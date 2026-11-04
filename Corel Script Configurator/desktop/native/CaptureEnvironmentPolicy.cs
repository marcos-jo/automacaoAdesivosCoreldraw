using System;
using Microsoft.Win32;

namespace FourPrintShops.PrintConfig {
    internal static class CaptureEnvironmentPolicy {
        internal static bool ManualDefault(object value) { return value is int && (int)value == 1; }
        internal static void RequireManualDefault() {
            using(var settings=Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Windows",false)) {
                if(settings==null||!ManualDefault(settings.GetValue("LegacyDefaultPrinterMode"))) {
                    throw new InvalidOperationException("Captura bloqueada. Desative a gestão automática da impressora padrão nas configurações do Windows antes do ensaio FILE:. O aplicativo não altera essa preferência.");
                }
            }
        }
    }
}
