using System;

namespace FourPrintShops.PrintConfig {
    // O conteúdo privado nunca é serializado. Só o driver pode interpretá-lo.
    internal sealed class DevModeState : IDisposable {
        private byte[] bytes;
        public int Revision { get; private set; }
        public DevModeState(byte[] initial) { Validate(initial); bytes=(byte[])initial.Clone(); Revision=1; }
        public static int Validate(byte[] value) {
            if(value==null || value.Length<102 || value.Length>4*1024*1024) throw new InvalidOperationException("DEVMODE inválido.");
            int size=BitConverter.ToUInt16(value,68), extra=BitConverter.ToUInt16(value,70);
            if(size<102 || size+extra>value.Length) throw new InvalidOperationException("DEVMODE incompleto.");
            return size+extra;
        }
        public byte[] Copy() { if(bytes==null) throw new ObjectDisposedException("DevModeState"); return (byte[])bytes.Clone(); }
        public void Accept(byte[] candidate, int expectedRevision) {
            if(expectedRevision!=Revision) throw new InvalidOperationException("Configuração desatualizada.");
            Validate(candidate);
            byte[] replacement=(byte[])candidate.Clone();
            Array.Clear(bytes,0,bytes.Length); bytes=replacement; Revision++;
        }
        public void Dispose() { if(bytes!=null) Array.Clear(bytes,0,bytes.Length); bytes=null; }
    }
}
