using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using WindowsInput;

namespace WeiXinSchoolReport;

public static class AutoUI
{
    private static readonly InputSimulator Input = new();

    public static void ClickOnScreen(Point location)
    {
        Cursor.Position = location;
        Thread.Sleep(1000);
        Input.Mouse.LeftButtonClick();
    }

    public static Rectangle GetWindowRectangle(IntPtr windowHandle)
    {
        var rect = new Rect();
        GetWindowRect(windowHandle, ref rect);
        return new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left + 1, rect.Bottom - rect.Top + 1);
    }

    public static void ClickOnWindow(IntPtr windowHandle, Point location)
    {
        var rect = GetWindowRectangle(windowHandle);
        ClickOnScreen(new Point(rect.X + location.X, rect.Y + location.Y));
    }

    public static void ScrollToEnd(int click, int testHeight = 500)
    {
        var cursorPosition = Cursor.Position;
        var captureRectangle = new Rectangle(cursorPosition, new Size(1, testHeight));
        var oldBitmap = CaptureScreen(captureRectangle);

        while (true)
        {
            Input.Mouse.VerticalScroll(click);
            Thread.Sleep(500);

            var newBitmap = CaptureScreen(captureRectangle);

            if (BitmapHelper.IsSame(oldBitmap, newBitmap))
                break;

            oldBitmap = newBitmap;
        }
    }

    public static bool Scroll(int click, int testHeight = 500)
    {
        var cursorPosition = Cursor.Position;
        var captureRectangle = new Rectangle(cursorPosition, new Size(1, testHeight));
        var oldBitmap = CaptureScreen(captureRectangle);

        Input.Mouse.VerticalScroll(click);
        Thread.Sleep(500);

        var newBitmap = CaptureScreen(captureRectangle);
        return !BitmapHelper.IsSame(oldBitmap, newBitmap);
    }

    public static Bitmap CaptureScreen(Rectangle rectangle)
    {
        var captureBitmap = new Bitmap(rectangle.Width, rectangle.Height, PixelFormat.Format32bppArgb);
        var captureGraphics = Graphics.FromImage(captureBitmap);
        captureGraphics.CopyFromScreen(rectangle.Location, Point.Empty, rectangle.Size);
        Directory.CreateDirectory("CaptureLogs");
        captureBitmap.Save($@"CaptureLogs\{rectangle.Left}, {rectangle.Top}, {rectangle.Width}, {rectangle.Height}.png");
        return captureBitmap;
    }

    public static Point? FindBitmapInScreen(Bitmap bitmap, Rectangle rectangleOnScreen, bool bottomUp)
    {
        var screenBitmap = CaptureScreen(rectangleOnScreen);
        var result = BitmapHelper.FindSubImage(screenBitmap, bitmap, bottomUp);
        if (!result.HasValue)
            return null;

        return rectangleOnScreen.Location + new Size(result.Value) + bitmap.Size / 2;
    }

    public static Point? FindBitmapInWindow(Bitmap bitmap, IntPtr windowHandle, bool bottomUp)
    {
        var windowRectangle = GetWindowRectangle(windowHandle);
        var result = FindBitmapInScreen(bitmap, windowRectangle, bottomUp);
        if (!result.HasValue)
            return null;

        return result.Value - new Size(windowRectangle.Location);
    }

    public static int FindColorInWindowColumn(IntPtr windowHandle, int x, Color color)
    {
        var windowRectangle = GetWindowRectangle(windowHandle);

        var bitmap = CaptureScreen(new Rectangle(
            windowRectangle.Location + new Size(x, 0), new Size(1, windowRectangle.Height)));
        var matchingCount = 0;

        for (var y = windowRectangle.Height - 1; y >= 0; y--)
            if (bitmap.GetPixel(0, y) == color)
            {
                matchingCount++;
                if (matchingCount < 5)
                    continue;

                return y;
            }
            else
            {
                matchingCount = 0;
            }

        return -1;
    }

    [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string? classname, string? captionName);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public readonly int Left; // x position of upper-left corner
        public readonly int Top; // y position of upper-left corner
        public readonly int Right; // x position of lower-right corner
        public readonly int Bottom; // y position of lower-right corner
    }

    [DllImport("user32.dll")]
    private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, UIntPtr dwExtraInfo);

    private const int MOUSEEVENTF_MOVE = 0x0001;

    public static void WakeUpMonitor()
    {
        mouse_event(MOUSEEVENTF_MOVE, 0, 1, 0, UIntPtr.Zero);
        Thread.Sleep(40);
        mouse_event(MOUSEEVENTF_MOVE, 0, -1, 0, UIntPtr.Zero);
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

    private const int WM_SYSCOMMAND = 0x112;
    private const int SC_MONITORPOWER = 0xF170;

    public static void TurnOffMonitor()
    {
        SendMessage((IntPtr)0xFFFF, WM_SYSCOMMAND, (IntPtr)SC_MONITORPOWER, (IntPtr)2);
    }

    [DllImport("gdi32.dll")]
    private static extern unsafe bool SetDeviceGammaRamp(int hdc, void* ramp);

    private static bool _hdcInitialized;
    private static int _hdc;

    private static void InitializeClass()
    {
        if (_hdcInitialized)
            return;

        //Get the hardware device context of the screen, we can do
        //this by getting the graphics object of null (IntPtr.Zero)
        //then getting the HDC and converting that to an Int32.
        _hdc = Graphics.FromHwnd(IntPtr.Zero).GetHdc().ToInt32();

        _hdcInitialized = true;
    }

    /// <summary>
    ///     Set screen brightness
    /// </summary>
    /// <param name="brightness">0-255</param>
    /// <returns></returns>
    public static unsafe bool SetScreeGamma(short brightness)
    {
        InitializeClass();

        if (brightness > 255)
            brightness = 255;

        if (brightness < 0)
            brightness = 0;

        var gArray = stackalloc short[3 * 256];
        var idx = gArray;

        for (var j = 0; j < 3; j++)
        {
            for (var i = 0; i < 256; i++)
            {
                var arrayVal = i * (brightness + 128);

                if (arrayVal > 65535)
                    arrayVal = 65535;

                *idx = (short)arrayVal;
                idx++;
            }
        }

        //For some reason, this always returns false?
        var retVal = SetDeviceGammaRamp(_hdc, gArray);

        //Memory allocated through stackalloc is automatically free'd
        //by the CLR.

        return retVal;
    }

    private static PhysicalMonitorBrightnessController? _monitorBrightness;

    /// <summary>
    /// All monitors' brightness, 0-100.
    /// </summary>
    public static int MonitorBrightness
    {
        get
        {
            _monitorBrightness ??= new PhysicalMonitorBrightnessController();
            return _monitorBrightness.Get();
        }
        set
        {
            _monitorBrightness ??= new PhysicalMonitorBrightnessController();
            _monitorBrightness.Set((uint)value);
        }
    }
}
