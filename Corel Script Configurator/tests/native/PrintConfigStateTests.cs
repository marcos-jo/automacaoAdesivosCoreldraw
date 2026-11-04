using System;
using FourPrintShops.PrintConfig;

internal static class PrintConfigStateTests {
    private static void Check(bool ok,string name){if(!ok)throw new Exception(name);}
    private static byte[] Fixture(){byte[] mode=new byte[220+8192];for(int n=0;n<mode.Length;n++)mode[n]=(byte)(n%251);Array.Copy(BitConverter.GetBytes((ushort)220),0,mode,68,2);Array.Copy(BitConverter.GetBytes((ushort)8192),0,mode,70,2);return mode;}
    private static bool Same(byte[] a,byte[] b){if(a.Length!=b.Length)return false;for(int n=0;n<a.Length;n++)if(a[n]!=b[n])return false;return true;}
    public static int Main(){
        byte[] original=Fixture();using(var state=new DevModeState(original)){
            Check(Same(original,state.Copy()),"Private data preserved initially");
            byte[] canceled=state.Copy();canceled[300]=17;Array.Clear(canceled,0,canceled.Length);
            Check(Same(original,state.Copy())&&state.Revision==1,"Cancel preserves every byte and revision");
            byte[] accepted=state.Copy();accepted[76]=2;accepted[400]=32;state.Accept(accepted,1);
            Check(Same(accepted,state.Copy())&&state.Revision==2,"Accepted native blob retained in full");
            accepted[400]=8;Check(state.Copy()[400]==32,"No aliases to candidate buffer");
            bool stale=false;try{state.Accept(original,1);}catch(InvalidOperationException){stale=true;}Check(stale&&state.Revision==2&&state.Copy()[400]==32,"Stale results cannot overwrite confirmed state");
            byte[] invalid=Fixture();invalid[68]=255;invalid[69]=255;
            bool rejected=false;try{state.Accept(invalid,2);}catch(InvalidOperationException){rejected=true;}Check(rejected&&state.Revision==2&&state.Copy()[400]==32,"Invalid buffer is transactional");
        }
        Console.WriteLine("PASSOU: 6 verificações de DEVMODE sintético, incluindo 8192 bytes privados, cancelamento, confirmação, isolamento de buffers, revisão e rejeição transacional.");return 0;
    }
}
