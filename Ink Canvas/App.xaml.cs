using InkCanvasPlus.Helpers;
using iNKORE.UI.WPF.Modern.Controls;
using System;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Security.Cryptography;
using System.Reflection;
using System.Windows;
using MessageBox = System.Windows.MessageBox;
using System.Threading;

namespace InkCanvasPlus
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        System.Threading.Mutex mutex;

        public static string[] StartArgs = null;
        public static string RootPath = Environment.GetEnvironmentVariable("APPDATA") + "\\Ink Canvas\\";

        public App()
        {
            this.Startup += new StartupEventHandler(App_Startup);
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            InkCanvasPlus.MainWindow.ShowNewMessage("抱歉，出现未预期的异常，可能导致 Ink Canvas 画板运行不稳定。\n建议保存墨迹后重启应用。", true);
            LogHelper.NewLog(e.Exception.ToString());
            e.Handled = true;
        }

        void App_Startup(object sender, StartupEventArgs e)
        {
            if (!StoreHelper.IsStoreApp) RootPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            LogHelper.NewLog(string.Format("Ink Canvas Starting (Version: {0})", Assembly.GetExecutingAssembly().GetName().Version.ToString()));

            string mutexName;
            using (var sha1 = SHA1.Create())
            {
                var hashBytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(AppDomain.CurrentDomain.BaseDirectory.ToLowerInvariant().Replace(":\\", ".").Replace("\\", ".")));
                mutexName = "top.khyan.InkCanvasPlus." + BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant().Substring(0, 12);
                LogHelper.NewLog($"Generated mutex name: {mutexName}");
            }

            mutex = new System.Threading.Mutex(true, mutexName, out bool ret);

            if (!ret && !e.Args.Contains("-m")) //-m multiple
            {
                LogHelper.NewLog("Detected existing instance");

                var ask = MessageBox.Show("Ink Canvas Plus 可能已经在运行了。请寻找屏幕上笑脸形状的图标；\n\n如果没找到，您可以点击「是」来重新启动它。", "Ink Canvas Plus 检测到其他实例", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (ask == MessageBoxResult.No)
                {
                    LogHelper.NewLog("User chose not to kill other instances. Exiting.");
                    Environment.Exit(0);
                }

                try
                {
                    var current = Process.GetCurrentProcess();
                    var others = Process.GetProcessesByName(current.ProcessName).Where(p => p.Id != current.Id).ToList();
                    foreach (var p in others)
                    {
                        try
                        {
                            LogHelper.NewLog($"Attempting to kill process {p.ProcessName} (PID {p.Id})");
                            p.Kill();
                            p.WaitForExit(2000);
                            LogHelper.NewLog($"Killed process {p.ProcessName} (PID {p.Id})");
                        }
                        catch (Exception ex)
                        {
                            LogHelper.NewLog($"Failed to kill process {p.ProcessName} (PID {p.Id}): {ex}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.NewLog(ex.ToString());
                }

                // Try to acquire the mutex again after attempting to kill other instances
                mutex = new System.Threading.Mutex(false, mutexName);
                bool got = false;
                try
                {
                    got = mutex.WaitOne(5000); // Wait up to 5 seconds to acquire the mutex
                }
                catch (AbandonedMutexException)
                {
                    got = true; // The mutex was abandoned, but we can still acquire it
                    LogHelper.NewLog("Mutex was abandoned, but acquired successfully.");
                }
                if (!got)
                {
                    MessageBox.Show("无法获得程序实例控制，程序将退出。", "Ink Canvas Plus 遇到错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    LogHelper.NewLog("Ink Canvas automatically closed - failed to acquire mutex after killing instances");
                    Environment.Exit(0);
                }
                else
                {
                    LogHelper.NewLog("Acquired mutex after killing other instances. Continuing.");
                }
            }

            StartArgs = e.Args;
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            try
            {
                if (System.Windows.Forms.SystemInformation.MouseWheelScrollLines == -1)
                    e.Handled = false;
                else
                    try
                    {
                        ScrollViewerEx SenderScrollViewer = (ScrollViewerEx)sender;
                        SenderScrollViewer.ScrollToVerticalOffset(SenderScrollViewer.VerticalOffset - e.Delta * 10 * System.Windows.Forms.SystemInformation.MouseWheelScrollLines / (double)120);
                        e.Handled = true;
                    }
                    catch
                    {
                    }
            }
            catch
            {
            }
        }
    }
}
