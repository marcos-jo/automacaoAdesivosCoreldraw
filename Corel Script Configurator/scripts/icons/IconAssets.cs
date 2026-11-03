using System;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Collections.Generic;

// Conversão determinística, sem redesenho, recorte ou fundo nos assets.
public static class IconAssets {
    public static readonly int[] Sizes = {16,24,32,48,64,128,256};
    public static Bitmap Fit(Image source, int size) {
        var bitmap = new Bitmap(size,size,PixelFormat.Format32bppArgb);
        bitmap.SetResolution(96,96);
        using(var g=Graphics.FromImage(bitmap))
        using(var attributes=new ImageAttributes()) {
            g.Clear(Color.Transparent);
            g.CompositingMode=CompositingMode.SourceCopy;
            g.CompositingQuality=CompositingQuality.HighQuality;
            g.InterpolationMode=InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode=PixelOffsetMode.HighQuality;
            attributes.SetWrapMode(WrapMode.TileFlipXY);
            float scale=Math.Min((float)size/source.Width,(float)size/source.Height);
            float width=source.Width*scale,height=source.Height*scale;
            float x=(size-width)/2,y=(size-height)/2;
            g.DrawImage(source,new PointF[]{new PointF(x,y),new PointF(x+width,y),new PointF(x,y+height)},new RectangleF(0,0,source.Width,source.Height),GraphicsUnit.Pixel,attributes);
        }
        return bitmap;
    }
    static byte[] Png(Bitmap bitmap) {
        using(var stream=new MemoryStream()){bitmap.Save(stream,ImageFormat.Png);return stream.ToArray();}
    }
    static byte[] Dib(Bitmap bitmap) {
        int size=bitmap.Width,maskStride=((size+31)/32)*4;
        using(var stream=new MemoryStream())
        using(var writer=new BinaryWriter(stream)) {
            writer.Write(40);writer.Write(size);writer.Write(size*2);
            writer.Write((ushort)1);writer.Write((ushort)32);writer.Write(0);
            writer.Write(size*size*4+maskStride*size);
            writer.Write(0);writer.Write(0);writer.Write(0);writer.Write(0);
            for(int y=size-1;y>=0;y--)for(int x=0;x<size;x++) {
                var c=bitmap.GetPixel(x,y);writer.Write(c.B);writer.Write(c.G);writer.Write(c.R);writer.Write(c.A);
            }
            for(int y=size-1;y>=0;y--) {
                var mask=new byte[maskStride];
                for(int x=0;x<size;x++)if(bitmap.GetPixel(x,y).A==0)mask[x/8]|=(byte)(0x80>>(x%8));
                writer.Write(mask);
            }
            return stream.ToArray();
        }
    }
    public static void Generate(string sourceFile,string outputBase) {
        using(var source=Image.FromFile(sourceFile,true)) {
            using(var large=Fit(source,1024))large.Save(outputBase+".png",ImageFormat.Png);
            var frames=new List<byte[]>();
            foreach(int size in Sizes)using(var small=Fit(source,size))frames.Add(size==256?Png(small):Dib(small));
            using(var stream=File.Create(outputBase+".ico"))
            using(var writer=new BinaryWriter(stream)) {
                writer.Write((ushort)0);writer.Write((ushort)1);writer.Write((ushort)Sizes.Length);
                int offset=6+16*Sizes.Length;
                for(int i=0;i<Sizes.Length;i++) {
                    int size=Sizes[i];writer.Write((byte)(size==256?0:size));writer.Write((byte)(size==256?0:size));
                    writer.Write((byte)0);writer.Write((byte)0);writer.Write((ushort)1);writer.Write((ushort)32);
                    writer.Write(frames[i].Length);writer.Write(offset);offset+=frames[i].Length;
                }
                foreach(var frame in frames)writer.Write(frame);
            }
        }
    }
    public static void Preview(string appFile,string pdfFile,string destination) {
        using(var bitmap=new Bitmap(900,700,PixelFormat.Format32bppArgb))
        using(var g=Graphics.FromImage(bitmap))
        using(var font=new Font("Segoe UI",10))
        using(var app=Image.FromFile(appFile))
        using(var pdf=Image.FromFile(pdfFile)) {
            g.Clear(Color.White);
            for(int row=0;row<4;row++) {
                bool dark=row%2==1;
                using(var background=new SolidBrush(dark?Color.FromArgb(40,40,40):Color.White))g.FillRectangle(background,0,row*175,900,175);
                using(var foreground=new SolidBrush(dark?Color.White:Color.Black)) {
                    g.DrawString((row<2?"Aplicativo":"PDF")+" - "+(dark?"fundo escuro":"fundo claro")+" (tamanho real)",font,foreground,8,row*175+5);
                    for(int i=0;i<6;i++) {
                        int size=Sizes[i],x=20+i*145;
                        using(var icon=Fit(row<2?app:pdf,size))g.DrawImageUnscaled(icon,x,row*175+32);
                        g.DrawString(size+" px",font,foreground,x,row*175+151);
                    }
                }
            }
            bitmap.Save(destination,ImageFormat.Png);
        }
    }
}
