using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WeiXinSchoolReport;

static class Program
{
    /// <summary>
    ///     The main entry point for the application.
    /// </summary>
    [STAThread]
    private static void Main()
    {
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

            Console.WriteLine(@"点开微信窗口");
            var screenBounds = Screen.PrimaryScreen.Bounds;
            Point? weChatIconPos = null;

            for (var i = 0; i < 10; i++)
            {
                weChatIconPos = AutoUI.FindBitmapInScreen(
                    BitmapHelper.FromFile("微信图标.png")!, new Rectangle(0, screenBounds.Height - 230, 170, 200), false);
                if (weChatIconPos != null)
                    break;
                Thread.Sleep(500);
            }

            if (weChatIconPos == null)
            {
                MessageBox.Show(@"找不到微信小图标");
                return;
            }

            AutoUI.ClickOnScreen(weChatIconPos.Value);
            Thread.Sleep(2000);

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
                return;
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
                return;
            }

            Console.WriteLine(@"点击最后一个“学生每日健康打卡”");
            AutoUI.ClickOnWindow(weChatWnd, lastMsg.Value);
            Console.WriteLine(@"等待答题窗口刷新");
            Thread.Sleep(10000);

            Console.WriteLine(@"查找弹出的健康打卡窗口");
            var viewWnd = AutoUI.FindWindow("CefWebViewWnd", "微信");
            var viewRectangle = AutoUI.GetWindowRectangle(viewWnd);

            Console.WriteLine(@"查找“上报已提交，直接退出”");
            if (AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("上报已提交.png")!, viewWnd, false) != null)
            {
                Console.WriteLine(@"上报已提交，直接退出");
                Thread.Sleep(3000);
                return;
            }

            Console.WriteLine(@"查找“题目内容有更新，请确认后填写”");
            var iKnownButton = AutoUI.FindBitmapInWindow(BitmapHelper.FromFile("我知道了.png")!, viewWnd, false);
            if (iKnownButton != null)
            {
                Console.WriteLine(@"发现“题目内容有更新，请确认后填写”");
                Console.WriteLine(@"点击“我知道了”");
                AutoUI.ClickOnWindow(viewWnd, iKnownButton.Value);
                Thread.Sleep(1000);

                Console.WriteLine(@"尝试自动答题");
                while (true)
                {
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
            else
            {
                Console.WriteLine(@"滚动到最下方");
                Cursor.Position = viewRectangle.Location + viewRectangle.Size / 2;
                AutoUI.ScrollToEnd(-100);
            }

            Console.WriteLine(@"查找“同意并提交”");
            var commitButton = AutoUI.FindBitmapInWindow(
                BitmapHelper.FromFile("同意并提交.png")!, viewWnd, true);
            if (commitButton == null)
            {
                MessageBox.Show(@"找不到按钮“同意并提交”");
                return;
            }

            Console.WriteLine(@"点击“同意并提交”");
            AutoUI.ClickOnWindow(viewWnd, commitButton.Value);
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
}
