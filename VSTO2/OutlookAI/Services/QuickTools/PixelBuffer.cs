using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>24-bit pixels in B, G, R order, rows packed without padding.</summary>
    internal sealed class PixelBuffer
    {
        public PixelBuffer(int width, int height)
        {
            if (width <= 0 || height <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            Width = width;
            Height = height;
            Data = new byte[checked(width * height * 3)];
        }

        public int Width { get; }
        public int Height { get; }
        public byte[] Data { get; }

        public static PixelBuffer FromBitmap(Bitmap bitmap)
        {
            var buffer = new PixelBuffer(bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            try
            {
                var row = buffer.Width * 3;
                for (var y = 0; y < buffer.Height; y++)
                {
                    Marshal.Copy(data.Scan0 + y * data.Stride, buffer.Data, y * row, row);
                }
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
            return buffer;
        }

        public Bitmap ToBitmap()
        {
            var bitmap = new Bitmap(Width, Height, PixelFormat.Format24bppRgb);
            try
            {
                var data = bitmap.LockBits(new Rectangle(0, 0, Width, Height), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
                try
                {
                    var row = Width * 3;
                    for (var y = 0; y < Height; y++)
                    {
                        Marshal.Copy(Data, y * row, data.Scan0 + y * data.Stride, row);
                    }
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }
                return bitmap;
            }
            catch
            {
                bitmap.Dispose();
                throw;
            }
        }

        public PixelBuffer Clone()
        {
            var copy = new PixelBuffer(Width, Height);
            Buffer.BlockCopy(Data, 0, copy.Data, 0, Data.Length);
            return copy;
        }
    }
}
