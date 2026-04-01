using AutoUpdaterDotNET;
using InkCanvasPlus.Helpers;
using iNKORE.UI.WPF.Modern;
using iNKORE.UI.WPF.Modern.Helpers;
using IWshRuntimeLibrary;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Input.StylusPlugIns;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Application = System.Windows.Application;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Point = System.Windows.Point;
using Timer = System.Timers.Timer;

namespace InkCanvasPlus
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint VK_D = 0x44;
        private const int HOTKEY_ID = 9000;
        private const int WM_HOTKEY = 0x0312;

        #region Window Initialization

        public MainWindow()
        {
            InitializeComponent();

            BorderSettings.Opacity = 0;
            BorderSettings.Visibility = Visibility.Collapsed;
            StackPanelToolButtons.Visibility = Visibility.Collapsed;
            BorderDrawShape.Visibility = Visibility.Collapsed;
            GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;

            if (App.StartArgs.Contains("-b")) //-b border
            {
                AllowsTransparency = false;
                WindowStyle = WindowStyle.SingleBorderWindow;
                ResizeMode = ResizeMode.CanResize;
                Background = new SolidColorBrush(StringToColor("#FFF2F2F2"));
                Topmost = false;
            }

            if (!App.StartArgs.Contains("-o")) //-old ui
            {
                GroupBoxAppearance.Visibility = Visibility.Collapsed;
                ViewBoxStackPanelMain.Visibility = Visibility.Collapsed;
                ViewBoxStackPanelShapes.Visibility = Visibility.Collapsed;
                HideSubPanels();

                ViewboxFloatingBar.Margin = new Thickness(SystemParameters.WorkArea.Left + 80, SystemParameters.WorkArea.Top + SystemParameters.WorkArea.Height - 80, -2000, -200);
                if (Settings.Appearance.IsAutoCollapseFloatBar)
                {
                    new Thread(new ThreadStart(() =>
                    {
                        Thread.Sleep(3000);
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (Main_Grid.Background == Brushes.Transparent)
                            {
                                SetBorderFloatingBarMainControlsVisibility(false);
                            }
                        });
                    })).Start();
                }

            }
            else
            {
                GroupBoxAppearanceNewUI.Visibility = Visibility.Collapsed;
                ViewboxFloatingBar.Visibility = Visibility.Collapsed;
                GridForRecoverOldUI.Visibility = Visibility.Collapsed;
            }

            if (File.Exists("debug.ini")) Label.Visibility = Visibility.Visible;

            InitTimers();
            timeMachine.OnRedoStateChanged += TimeMachine_OnRedoStateChanged;
            timeMachine.OnUndoStateChanged += TimeMachine_OnUndoStateChanged;
            inkCanvas.Strokes.StrokesChanged += StrokesOnStrokesChanged;

            Microsoft.Win32.SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;

            AutoUpdater.RunUpdateAsAdmin = false;
            AutoUpdater.TopMost = true;
            AutoUpdater.ApplicationExitEvent += () =>
            {
                Environment.Exit(0);
            };
            CheckForUpdate();

            UpdateWindowTitle();
        }

        private void UpdateWindowTitle()
        {
            string title = "Ink Canvas Plus - ";
            if (currentMode == 0)
            {
                if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                {
                    title += "Presentation";
                }
                else
                {
                    title += "Desktop";
                }
            }
            else if (currentMode == 1)
            {
                title += "Board";
            }
            title += " ";
            if (Main_Grid.Background == Brushes.Transparent)
            {
                title += "Idle";
            }
            else
            {
                title += "Active";
            }
            this.Title = title;
        }

        #endregion

        #region Timer

        Timer timerCheckPPT = new Timer();
        Timer timerKillProcess = new Timer();

        private void InitTimers()
        {
            timerCheckPPT.Elapsed += TimerCheckPPT_Elapsed;
            timerCheckPPT.Interval = 1000;

            timerKillProcess.Elapsed += TimerKillProcess_Elapsed;
            timerKillProcess.Interval = 1000;
        }

        private void TimerKillProcess_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                // 希沃相关： easinote swenserver RemoteProcess EasiNote.MediaHttpService smartnote.cloud EasiUpdate smartnote EasiUpdate3 EasiUpdate3Protect SeewoP2P CefSharp.BrowserSubprocess SeewoUploadService
                string arg = "/F";
                if (Settings.Automation.IsAutoKillPptService)
                {
                    Process[] processes = Process.GetProcessesByName("PPTService");
                    if (processes.Length > 0)
                    {
                        arg += " /IM PPTService.exe";
                    }
                    processes = Process.GetProcessesByName("SeewoIwbAssistant");
                    if (processes.Length > 0)
                    {
                        arg += " /IM SeewoIwbAssistant.exe" +
                            " /IM Sia.Guard.exe";
                    }
                }
                if (arg != "/F")
                {
                    Process p = new Process();
                    p.StartInfo = new ProcessStartInfo("taskkill", arg);
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                }
                if (Settings.Automation.IsAutoKillEasiNote)
                {
                    Process[] processes = Process.GetProcessesByName("EasiNote");
                    if (processes.Length > 0)
                    {
                        AutoKillHelper.KillEasiNoteFloatBall();
                    }
                }
            }
            catch { }
        }

        #endregion Timer

    }

    #region Test for pen
    // A StylusPlugin that renders ink with a linear gradient brush effect.
    class CustomDynamicRenderer : DynamicRenderer
    {
        [ThreadStatic]
        static private Brush brush = null;

        [ThreadStatic]
        static private Pen pen = null;

        private Point prevPoint;

        protected override void OnStylusDown(RawStylusInput rawStylusInput)
        {
            // Allocate memory to store the previous point to draw from.
            prevPoint = new Point(double.NegativeInfinity, double.NegativeInfinity);
            base.OnStylusDown(rawStylusInput);
        }
        //protected override void OnDraw(System.Windows.Media.DrawingContext drawingContext, System.Windows.Input.StylusPointCollection stylusPoints, System.Windows.Media.Geometry geometry, System.Windows.Media.Brush fillBrush)
        //{


        //    ImageSource img = new BitmapImage(new Uri("pack://application:,,,/Resources/maobi.png"));

        //    //前一个点的绘制。
        //    Point prevPoint = new Point(double.NegativeInfinity,
        //                                double.NegativeInfinity);


        //    var w = Global.StrokeWidth + 15;    //输出时笔刷的实际大小


        //    Point pt = new Point(0, 0);
        //    Vector v = new Vector();            //前一个点与当前点的距离
        //    var subtractY = 0d;                 //当前点处前一点的Y偏移
        //    var subtractX = 0d;                 //当前点处前一点的X偏移
        //    var pointWidth = Global.StrokeWidth;
        //    double x = 0, y = 0;
        //    for (int i = 0; i < stylusPoints.Count; i++)
        //    {
        //        pt = (Point)stylusPoints[i];
        //        v = Point.Subtract(prevPoint, pt);

        //        Debug.WriteLine("X " + pt.X + "\t" + pt.Y);

        //        subtractY = (pt.Y - prevPoint.Y) / v.Length;    //设置stylusPoints两个点之间需要填充的XY偏移
        //        subtractX = (pt.X - prevPoint.X) / v.Length;

        //        if (w - v.Length < Global.StrokeWidth)          //控制笔刷大小
        //        {
        //            pointWidth = Global.StrokeWidth;
        //        }
        //        else
        //        {
        //            pointWidth = w - v.Length;                  //在两个点距离越大的时候，笔刷所展示的大小越小
        //        }


        //        for (double j = 0; j < v.Length; j = j + 1d)    //填充stylusPoints两个点之间
        //        {
        //            x = 0; y = 0;

        //            if (prevPoint.X == double.NegativeInfinity || prevPoint.Y == double.NegativeInfinity || double.PositiveInfinity == prevPoint.X || double.PositiveInfinity == prevPoint.Y)
        //            {
        //                y = pt.Y;
        //                x = pt.X;
        //            }
        //            else
        //            {
        //                y = prevPoint.Y + subtractY;
        //                x = prevPoint.X + subtractX;
        //            }

        //            drawingContext.DrawImage(img, new Rect(x - pointWidth / 2, y - pointWidth / 2, pointWidth, pointWidth));    //在当前点画笔刷图片
        //            prevPoint = new Point(x, y);


        //            if (double.IsNegativeInfinity(v.Length) || double.IsPositiveInfinity(v.Length))
        //            { break; }
        //        }
        //    }
        //    stylusPoints = null;
        //}
        protected override void OnDraw(DrawingContext drawingContext,
                                       StylusPointCollection stylusPoints,
                                       Geometry geometry, Brush fillBrush)
        {
            // Create a new Brush, if necessary.
            //brush ??= new LinearGradientBrush(Colors.Red, Colors.Blue, 20d);

            // Create a new Pen, if necessary.
            //pen ??= new Pen(brush, 2d);

            // Draw linear gradient ellipses between 
            // all the StylusPoints that have come in.
            for (int i = 0; i < stylusPoints.Count; i++)
            {
                Point pt = (Point)stylusPoints[i];
                Vector v = Point.Subtract(prevPoint, pt);

                // Only draw if we are at least 4 units away 
                // from the end of the last ellipse. Otherwise, 
                // we're just redrawing and wasting cycles.
                if (v.Length > 4)
                {
                    // Set the thickness of the stroke based 
                    // on how hard the user pressed.
                    double radius = stylusPoints[i].PressureFactor * 10d;
                    drawingContext.DrawEllipse(brush, pen, pt, radius, radius);
                    prevPoint = pt;
                }
            }
        }
    }
    public class Global
    {
        public static double StrokeWidth = 2.5;
    }
    public class CustomRenderingInkCanvas : InkCanvas
    {
        CustomDynamicRenderer customRenderer = new CustomDynamicRenderer();

        public CustomRenderingInkCanvas() : base()
        {
            // Use the custom dynamic renderer on the
            // custom InkCanvas.
            this.DynamicRenderer = customRenderer;
        }

        protected override void OnStrokeCollected(InkCanvasStrokeCollectedEventArgs e)
        {
            // Remove the original stroke and add a custom stroke.
            this.Strokes.Remove(e.Stroke);
            CustomStroke customStroke = new CustomStroke(e.Stroke.StylusPoints);
            this.Strokes.Add(customStroke);

            // Pass the custom stroke to base class' OnStrokeCollected method.
            InkCanvasStrokeCollectedEventArgs args =
                new InkCanvasStrokeCollectedEventArgs(customStroke);
            base.OnStrokeCollected(args);
        }
    }// A class for rendering custom strokes
    class CustomStroke : Stroke
    {
        Brush brush;
        Pen pen;

        public CustomStroke(StylusPointCollection stylusPoints)
            : base(stylusPoints)
        {
            // Create the Brush and Pen used for drawing.
            brush = new LinearGradientBrush(Colors.Red, Colors.Blue, 20d);
            pen = new Pen(brush, 2d);
        }
        //protected override void DrawCore(DrawingContext drawingContext, DrawingAttributes drawingAttributes)
        //{


        //            ImageSource img = new BitmapImage(new Uri("pack://application:,,,/Resources/maobi.png"));

        //    //前一个点的绘制。
        //    Point prevPoint = new Point(double.NegativeInfinity,
        //                                double.NegativeInfinity);


        //    var w = Global.StrokeWidth + 15;    //输出时笔刷的实际大小


        //    Point pt = new Point(0, 0);
        //    Vector v = new Vector();            //前一个点与当前点的距离
        //    var subtractY = 0d;                 //当前点处前一点的Y偏移
        //    var subtractX = 0d;                 //当前点处前一点的X偏移
        //    var pointWidth = Global.StrokeWidth;
        //    double x = 0, y = 0;
        //    for (int i = 0; i < stylusPoints.Count; i++)
        //    {
        //        pt = (Point)stylusPoints[i];
        //        v = Point.Subtract(prevPoint, pt);

        //        Debug.WriteLine("X " + pt.X + "\t" + pt.Y);

        //        subtractY = (pt.Y - prevPoint.Y) / v.Length;    //设置stylusPoints两个点之间需要填充的XY偏移
        //        subtractX = (pt.X - prevPoint.X) / v.Length;

        //        if (w - v.Length < Global.StrokeWidth)          //控制笔刷大小
        //        {
        //            pointWidth = Global.StrokeWidth;
        //        }
        //        else
        //        {
        //            pointWidth = w - v.Length;                  //在两个点距离越大的时候，笔刷所展示的大小越小
        //        }


        //        for (double j = 0; j < v.Length; j = j + 1d)    //填充stylusPoints两个点之间
        //        {
        //            x = 0; y = 0;

        //            if (prevPoint.X == double.NegativeInfinity || prevPoint.Y == double.NegativeInfinity || double.PositiveInfinity == prevPoint.X || double.PositiveInfinity == prevPoint.Y)
        //            {
        //                y = pt.Y;
        //                x = pt.X;
        //            }
        //            else
        //            {
        //                y = prevPoint.Y + subtractY;
        //                x = prevPoint.X + subtractX;
        //            }

        //            drawingContext.DrawImage(img, new Rect(x - pointWidth / 2, y - pointWidth / 2, pointWidth, pointWidth));    //在当前点画笔刷图片
        //            prevPoint = new Point(x, y);


        //            if (double.IsNegativeInfinity(v.Length) || double.IsPositiveInfinity(v.Length))
        //            { break; }
        //        }
        //    }
        //    stylusPoints = null;
        //}
        protected override void DrawCore(DrawingContext drawingContext,
                                         DrawingAttributes drawingAttributes)
        {
            // Allocate memory to store the previous point to draw from.
            Point prevPoint = new Point(double.NegativeInfinity,
                                        double.NegativeInfinity);

            // Draw linear gradient ellipses between
            // all the StylusPoints in the Stroke.
            for (int i = 0; i < this.StylusPoints.Count; i++)
            {
                Point pt = (Point)this.StylusPoints[i];
                Vector v = Point.Subtract(prevPoint, pt);

                // Only draw if we are at least 4 units away
                // from the end of the last ellipse. Otherwise,
                // we're just redrawing and wasting cycles.
                if (v.Length > 4)
                {
                    // Set the thickness of the stroke
                    // based on how hard the user pressed.
                    double radius = this.StylusPoints[i].PressureFactor * 10d;
                    drawingContext.DrawEllipse(brush, pen, pt, radius, radius);
                    prevPoint = pt;
                }
            }
        }
    }
    #endregion
}
