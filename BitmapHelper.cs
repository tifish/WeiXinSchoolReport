using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace WeiXinSchoolReport;

public static class BitmapHelper
{
    [DllImport("msvcrt.dll")]
    private static extern int memcmp(IntPtr b1, IntPtr b2, long count);

    public static bool IsSame(Bitmap b1, Bitmap b2)
    {
        if (b1.Size != b2.Size)
            return false;

        var bd1 = b1.LockBits(new Rectangle(new Point(0, 0), b1.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        var bd2 = b2.LockBits(new Rectangle(new Point(0, 0), b2.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        try
        {
            var bd1Scan0 = bd1.Scan0;
            var bd2Scan0 = bd2.Scan0;

            var stride = bd1.Stride;
            var len = stride * b1.Height;

            return memcmp(bd1Scan0, bd2Scan0, len) == 0;
        }
        finally
        {
            b1.UnlockBits(bd1);
            b2.UnlockBits(bd2);
        }
    }

    public static Point? FindSubImage(Bitmap fullImage, Bitmap subImage, bool bottomUp = false)
    {
        if (fullImage.Width < subImage.Width || fullImage.Height < subImage.Height)
            return null;

        var pixelSize = Image.GetPixelFormatSize(PixelFormat.Format32bppArgb) / 8;

        var fullBits = fullImage.LockBits(new Rectangle(new Point(0, 0), fullImage.Size),
            ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        var subBits = subImage.LockBits(new Rectangle(new Point(0, 0), subImage.Size),
            ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            if (bottomUp)
                for (var y = fullBits.Height - subBits.Height; y >= 0; y--)
                {
                    if (HasSubImageInLine(y, out var position))
                        return position;
                }
            else
                for (var y = 0; y <= fullBits.Height - subBits.Height; y++)
                    if (HasSubImageInLine(y, out var position))
                        return position;

            return null;
        }
        finally
        {
            fullImage.UnlockBits(fullBits);
            subImage.UnlockBits(subBits);
        }

        bool HasSubImageInLine(int y, out Point position)
        {
            for (var x = 0; x <= fullBits.Width - subBits.Width; x++)
                // Found first line
                if (memcmp(fullBits.Scan0 + y * fullBits.Stride + x * pixelSize, subBits.Scan0, subBits.Stride)
                    == 0)
                    if (HasSubImageInPosition(new Point(x, y)))
                    {
                        position = new Point(x, y);
                        return true;
                    }

            position = Point.Empty;
            return false;
        }

        bool HasSubImageInPosition(Point position)
        {
            if (position.Y > fullBits.Height - subBits.Height)
                return false;

            for (var y = 1; y < subBits.Height; y++)
                if (memcmp(fullBits.Scan0 + (position.Y + y) * fullBits.Stride + position.X * pixelSize,
                        subBits.Scan0 + y * subBits.Stride,
                        subBits.Stride) != 0)
                    return false;

            return true;
        }
    }

    public static Bitmap? FromFile(string filePath)
    {
        if (!File.Exists(filePath))
            return null;

        return new Bitmap(Image.FromFile(filePath));
    }
}
