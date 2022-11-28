using System.Diagnostics;
using System.DirectoryServices.ActiveDirectory;
using System.Reflection;

namespace WeiXinSchoolReport;

static class Program
{
    /// <summary>
    ///     The main entry point for the application.
    /// </summary>
    [STAThread]
    private static void Main()
    {
        Environment.CurrentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;

        var originalBrightness = AutoUI.MonitorBrightness;
        AutoUI.MonitorBrightness = 80;

        if (!Debugger.IsAttached)
        {
            AutoUI.MonitorBrightness = 0;
            AutoUI.SetScreeGamma(0);
        }

        try
        {
            if (!Debugger.IsAttached)
            {
                AutoUI.WakeUpMonitor();
                Thread.Sleep(30000);
            }

            if (!OpenWeChat())
                return;

            if (!OpenPunchCardView())
                return;

            if (!PunchCard())
                return;

            Thread.Sleep(3000);
        }
        finally
        {
            if (!Debugger.IsAttached)
            {
                AutoUI.TurnOffMonitor();
                AutoUI.MonitorBrightness = originalBrightness;
                AutoUI.SetScreeGamma(128);
            }
        }
    }

    private static bool OpenWeChat()
    {
        Console.WriteLine(@"点开微信窗口");
        var screenBounds = Screen.PrimaryScreen.Bounds;
        var weChatIconPos = GetWeChatIconPos();

        if (weChatIconPos == null)
        {
            Process.Start(@"C:\Library\Software\Net\Chat\WeChat\WeChat.exe");
            var loginWnd = IntPtr.Zero;
            for (var i = 0; i < 10; i++)
            {
                loginWnd = AutoUI.FindWindow("WeChatLoginWndForPC", null);
                if (loginWnd != IntPtr.Zero)
                    break;

                Thread.Sleep(500);
            }

            if (loginWnd == IntPtr.Zero)
            {
                MessageBox.Show(@"无法启动微信登录窗口");
                return false;
            }

            var loginButton = AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("进入微信.png")!, loginWnd, false);
            if (loginButton == null)
            {
                MessageBox.Show(@"找不到“进入微信”按钮");
                return false;
            }

            AutoUI.ClickOnWindow(loginWnd, loginButton.Value);
            Thread.Sleep(5000);
            weChatIconPos = GetWeChatIconPos();
        }

        if (weChatIconPos == null)
        {
            MessageBox.Show(@"找不到微信小图标");
            return false;
        }

        AutoUI.ClickOnScreen(weChatIconPos.Value);
        Thread.Sleep(2000);
        return true;

        Point? GetWeChatIconPos()
        {
            // 微信有消息时图标会闪烁，所以多尝试几次
            for (var i = 0; i < 10; i++)
            {
                weChatIconPos = AutoUI.FindBitmapInScreen(
                    BitmapHelper.FromFile("微信图标.png")!, new Rectangle(0, screenBounds.Height - 230, 170, 200), false);
                if (weChatIconPos != null)
                    return weChatIconPos;
                Thread.Sleep(500);
            }

            return null;
        }
    }

    private static bool OpenPunchCardView()
    {
        Console.WriteLine(@"找到微信窗口");
        var weChatWnd = AutoUI.FindWindow("WeChatMainWndForPC", "微信");

        // 滚动联系人到最上方
        var weChatRectangle = AutoUI.GetWindowRectangle(weChatWnd);
        Cursor.Position = weChatRectangle.Location + new Size(200, 100);
        Thread.Sleep(500);
        AutoUI.ScrollToEnd(100);

        Console.WriteLine(@"点击联系人“学校通知”");
        var schoolNotice = AutoUI.FindBitmapInWindow(
            BitmapHelper.FromFile("学校通知.png")!, weChatWnd, false);
        if (schoolNotice == null)
        {
            MessageBox.Show(@"找不到联系人“学校通知”");
            return false;
        }

        AutoUI.ClickOnWindow(weChatWnd, schoolNotice.Value);
        Thread.Sleep(1000);

        Console.WriteLine(@"查找学校");
        var school = AutoUI.FindBitmapInWindow(
            BitmapHelper.FromFile("学校.png")!, weChatWnd, false);
        if (school != null)
        {
            Console.WriteLine(@"点击学校");
            AutoUI.ClickOnWindow(weChatWnd, school.Value);
        }
        else
        {
            Console.WriteLine(@"图片有变化，导致找不到条目学校图标，尝试盲操");
            AutoUI.ClickOnWindow(weChatWnd, school ?? new Point(415, 195));
        }

        Thread.Sleep(1000);

        Console.WriteLine(@"查找“学生每日健康打卡”");
        var lastMsg = AutoUI.FindBitmapInWindow(
            BitmapHelper.FromFile("学生每日健康打卡.png")!, weChatWnd, true);
        if (lastMsg == null)
        {
            MessageBox.Show(@"找不到消息“学生每日健康打卡”");
            return false;
        }

        Console.WriteLine(@"点击最后一个“学生每日健康打卡”");
        AutoUI.ClickOnWindow(weChatWnd, lastMsg.Value);
        Console.WriteLine(@"等待答题窗口刷新");
        Thread.Sleep(10000);

        return true;
    }

    private static bool PunchCard()
    {
        Console.WriteLine(@"查找弹出的健康打卡窗口");
        var viewWnd = AutoUI.FindWindow("CefWebViewWnd", "微信");
        var viewRectangle = AutoUI.GetWindowRectangle(viewWnd);

        Console.WriteLine(@"查找“上报已提交”");
        if (AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("上报已提交.png")!, viewWnd, false) != null)
        {
            Console.WriteLine(@"上报已提交，直接退出");
            Thread.Sleep(3000);
            return true;
        }

        Console.WriteLine(@"查找“题目内容有更新，请确认后填写”");
        var iKnownButton = AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("我知道了.png")!, viewWnd, false);
        if (iKnownButton != null)
        {
            Console.WriteLine(@"发现“题目内容有更新，请确认后填写”");
            Console.WriteLine(@"点击“我知道了”");
            AutoUI.ClickOnWindow(viewWnd, iKnownButton.Value);
            Thread.Sleep(1000);
        }

        Console.WriteLine(@"不论需不需要重新答题，都尝试回答一遍");
        AutoAnswer(viewWnd, viewRectangle);

        Console.WriteLine(@"查找“同意并提交”");
        var commitButton = AutoUI.FindBitmapInWindow(
            BitmapHelper.FromFile("同意并提交.png")!, viewWnd, true);
        if (commitButton == null)
        {
            MessageBox.Show(@"找不到按钮“同意并提交”");
            return false;
        }

        Console.WriteLine(@"点击“同意并提交”");
        AutoUI.ClickOnWindow(viewWnd, commitButton.Value);

        return true;
    }

    private static void AutoAnswer(IntPtr viewWnd, Rectangle viewRectangle)
    {
        while (true)
        {
            var answerHeath = AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("答案健康.png")!, viewWnd, false);
            if (answerHeath != null)
            {
                Console.WriteLine(@"回答“健康”");
                AutoUI.ClickOnWindow(viewWnd, answerHeath.Value);
                continue;
            }

            var answerNo = AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("答案否.png")!, viewWnd, false);
            if (answerNo != null)
            {
                Console.WriteLine(@"回答“否”");
                AutoUI.ClickOnWindow(viewWnd, answerNo.Value);
                continue;
            }

            Console.WriteLine(@"向下滚动，继续答题。滚动到底则退出");
            Cursor.Position = viewRectangle.Location + viewRectangle.Size / 2;
            if (!AutoUI.Scroll(-3))
                break;
        }
    }
}
