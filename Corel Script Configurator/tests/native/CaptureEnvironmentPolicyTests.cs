using System;
using FourPrintShops.PrintConfig;

internal static class CaptureEnvironmentPolicyTests {
    public static int Main() {
        foreach(object value in new object[]{null,0,2,-1,"1",1L}) {
            if(CaptureEnvironmentPolicy.ManualDefault(value))throw new Exception("Estado automático/indefinido aceito.");
        }
        if(!CaptureEnvironmentPolicy.ManualDefault(1))throw new Exception("Modo manual recusado.");
        Console.WriteLine("PASSOU: 7 verificações da proteção da impressora padrão, sem alterações no Windows.");
        return 0;
    }
}
