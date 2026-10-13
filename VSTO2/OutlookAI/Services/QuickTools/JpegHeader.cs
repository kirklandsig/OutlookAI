namespace OutlookAI.Services.QuickTools
{
    /// <summary>Size and component count from a JPEG's frame header.</summary>
    internal static class JpegHeader
    {
        public static bool TryRead(byte[] jpeg, out int width, out int height, out int components)
        {
            width = height = components = 0;
            if (jpeg == null || jpeg.Length < 4 || jpeg[0] != 0xFF || jpeg[1] != 0xD8) return false;
            var i = 2;
            while (i + 3 < jpeg.Length)
            {
                if (jpeg[i] != 0xFF) { i++; continue; }
                var marker = jpeg[i + 1];
                if (marker == 0xFF) { i++; continue; }
                i += 2;
                if (marker == 0xD8 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7)) continue;
                if (marker == 0xD9 || marker == 0xDA) return false;
                if (i + 1 >= jpeg.Length) return false;
                var length = (jpeg[i] << 8) | jpeg[i + 1];
                if (length < 2 || i + length > jpeg.Length) return false;
                var isFrame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                if (isFrame)
                {
                    if (length < 8) return false;
                    height = (jpeg[i + 3] << 8) | jpeg[i + 4];
                    width = (jpeg[i + 5] << 8) | jpeg[i + 6];
                    components = jpeg[i + 7];
                    return width > 0 && height > 0;
                }
                i += length;
            }
            return false;
        }
    }
}
