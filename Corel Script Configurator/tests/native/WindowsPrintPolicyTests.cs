using System;
using FourPrintShops.PrintConfig;
internal static class WindowsPrintPolicyTests {
    private static void Check(bool value){if(!value)throw new Exception("Política de escala/geometria inválida");}
    public static int Main(){
        var mode=new byte[220+8192];Array.Copy(BitConverter.GetBytes((ushort)220),0,mode,68,2);Array.Copy(BitConverter.GetBytes((ushort)8192),0,mode,70,2);
        Check(WindowsPrintJob.EffectiveScale(120,mode)==120);
        Array.Copy(BitConverter.GetBytes(16u|256u),0,mode,72,4);Array.Copy(BitConverter.GetBytes((short)100),0,mode,84,2);
        Check(WindowsPrintJob.EffectiveScale(87.5,mode)==87.5);
        Array.Copy(BitConverter.GetBytes((short)120),0,mode,84,2);
        Check(WindowsPrintJob.EffectiveScale(120,mode)==100);
        Array.Copy(BitConverter.GetBytes((short)50),0,mode,84,2);Check(WindowsPrintJob.EffectiveScale(87.5,mode)==100);
        Check(WindowsPrintJob.Pixels(72,600,100)==600);Check(WindowsPrintJob.Pixels(72,600,120)==720);
        Check(WindowsPrintJob.Pixels(72,600,87.5)==525);Check(WindowsPrintJob.Pixels(1,72,1)==1);
        int rejected=0;foreach(double scale in new[]{0,401,Double.NaN,Double.PositiveInfinity})try{WindowsPrintJob.EffectiveScale(scale,mode);}catch(InvalidOperationException){rejected++;}Check(rejected==4);
        bool invalid=false;try{WindowsPrintJob.Pixels(Double.NaN,600,100);}catch(InvalidOperationException){invalid=true;}Check(invalid);
        Check(BitConverter.ToInt16(mode,84)==50&&BitConverter.ToUInt16(mode,70)==8192);
        Console.WriteLine("PASSOU: 11 verificações Windows de escala única, precedência do driver, decimais, geometria e DEVMODE preservado. Nenhuma API de envio executada.");return 0;
    }
}
