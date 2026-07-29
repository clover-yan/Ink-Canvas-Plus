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
        #region Touch Events

        #region Multi-Touch

        bool isInMultiTouchMode = false;
        private void BorderMultiTouchMode_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (isInMultiTouchMode)
            {
                inkCanvas.StylusDown -= MainWindow_StylusDown;
                inkCanvas.StylusMove -= MainWindow_StylusMove;
                inkCanvas.StylusUp -= MainWindow_StylusUp;
                inkCanvas.TouchDown -= MainWindow_TouchDown;
                inkCanvas.TouchDown += Main_Grid_TouchDown;
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                inkCanvas.Children.Clear();
                isInMultiTouchMode = false;
                SymbolIconMultiTouchMode.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.People;
            }
            else
            {
                inkCanvas.StylusDown += MainWindow_StylusDown;
                inkCanvas.StylusMove += MainWindow_StylusMove;
                inkCanvas.StylusUp += MainWindow_StylusUp;
                inkCanvas.TouchDown -= Main_Grid_TouchDown;
                inkCanvas.TouchDown += MainWindow_TouchDown;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.Children.Clear();
                isInMultiTouchMode = true;
                SymbolIconMultiTouchMode.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Contact;
            }
        }

        private void MainWindow_TouchDown(object sender, TouchEventArgs e)
        {
            double boundWidth = GetTouchBoundWidth(e);
            if (boundWidth > BoundsWidth)
            {
                inkCanvas.EraserShape = new EllipseStylusShape(boundWidth, boundWidth);
                TouchDownPointsList[e.TouchDevice.Id] = InkCanvasEditingMode.EraseByPoint;
                inkCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
            }
            else
            {
                TouchDownPointsList[e.TouchDevice.Id] = InkCanvasEditingMode.None;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
            }
        }

        private void MainWindow_StylusDown(object sender, StylusDownEventArgs e)
        {
            TouchDownPointsList[e.StylusDevice.Id] = InkCanvasEditingMode.None;
        }

        private void MainWindow_StylusUp(object sender, StylusEventArgs e)
        {
            try
            {
                inkCanvas.Strokes.Add(GetStrokeVisual(e.StylusDevice.Id).Stroke);
                inkCanvas.Children.Remove(GetVisualCanvas(e.StylusDevice.Id));

                inkCanvas_StrokeCollected(inkCanvas, new InkCanvasStrokeCollectedEventArgs(GetStrokeVisual(e.StylusDevice.Id).Stroke));
            }
            catch (Exception ex)
            {
                Label.Content = ex.ToString();
            }
            try
            {
                StrokeVisualList.Remove(e.StylusDevice.Id);
                VisualCanvasList.Remove(e.StylusDevice.Id);
                TouchDownPointsList.Remove(e.StylusDevice.Id);
                if (StrokeVisualList.Count == 0 || VisualCanvasList.Count == 0 || TouchDownPointsList.Count == 0)
                {
                    inkCanvas.Children.Clear();
                    StrokeVisualList.Clear();
                    VisualCanvasList.Clear();
                    TouchDownPointsList.Clear();
                }
            }
            catch { }
        }

        private void MainWindow_StylusMove(object sender, StylusEventArgs e)
        {
            try
            {
                if (GetTouchDownPointsList(e.StylusDevice.Id) != InkCanvasEditingMode.None) return;
                try
                {
                    if (e.StylusDevice.StylusButtons[1].StylusButtonState == StylusButtonState.Down) return;
                }
                catch { }
                var strokeVisual = GetStrokeVisual(e.StylusDevice.Id);
                var stylusPointCollection = e.GetStylusPoints(this);
                foreach (var stylusPoint in stylusPointCollection)
                {
                    strokeVisual.Add(new StylusPoint(stylusPoint.X, stylusPoint.Y, stylusPoint.PressureFactor));
                }

                strokeVisual.Redraw();
            }
            catch { }
        }

        private StrokeVisual GetStrokeVisual(int id)
        {
            if (StrokeVisualList.TryGetValue(id, out var visual))
            {
                return visual;
            }

            var strokeVisual = new StrokeVisual(inkCanvas.DefaultDrawingAttributes.Clone());
            StrokeVisualList[id] = strokeVisual;
            var visualCanvas = new VisualCanvas(strokeVisual);
            VisualCanvasList[id] = visualCanvas;
            inkCanvas.Children.Add(visualCanvas);

            return strokeVisual;
        }

        private VisualCanvas GetVisualCanvas(int id)
        {
            if (VisualCanvasList.TryGetValue(id, out var visualCanvas))
            {
                return visualCanvas;
            }
            return null;
        }

        private InkCanvasEditingMode GetTouchDownPointsList(int id)
        {
            if (TouchDownPointsList.TryGetValue(id, out var inkCanvasEditingMode))
            {
                return inkCanvasEditingMode;
            }
            return inkCanvas.EditingMode;
        }

        private Dictionary<int, InkCanvasEditingMode> TouchDownPointsList { get; } = new Dictionary<int, InkCanvasEditingMode>();
        private Dictionary<int, StrokeVisual> StrokeVisualList { get; } = new Dictionary<int, StrokeVisual>();
        private Dictionary<int, VisualCanvas> VisualCanvasList { get; } = new Dictionary<int, VisualCanvas>();

        #endregion

        int lastTouchDownTime = 0, lastTouchUpTime = 0;

        Point iniP = new Point(0, 0);
        bool isLastTouchEraser = false;
        private bool forcePointEraser = true;
        private bool _lockSmith = false; //临时停用双指手势
        private bool lockSmithPPT = false;
        private bool lockSmithDesktop = false;

        private void ChangeLockSmithState(bool state)
        {
            _lockSmith = state;
            if (_lockSmith) LockSmithSymbol.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.UnPin;
            else LockSmithSymbol.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Pin;

        }

        private double GetEraserSizeCoefficient()
        {
            double k = 1;
            switch (Settings.Canvas.EraserSize)
            {
                case -2:
                    k = 0.5;
                    break;
                case -1:
                    k = 0.8;
                    break;
                case 1:
                    k = 1.6;
                    break;
                case 2:
                    k = 2;
                    break;
                case 3:
                    k = 2.5;
                    break;
                case 4:
                    k = 3.6;
                    break;
            }
            return k;
        }

        private void Main_Grid_TouchDown(object sender, TouchEventArgs e)
        {
            BorderClearInDelete.Visibility = Visibility.Collapsed;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }

            if (NeedUpdateIniP())
            {
                iniP = e.GetTouchPoint(inkCanvas).Position;
            }
            if (drawingShapeMode == 9 && isFirstTouchCuboid == false)
            {
                MouseTouchMove(iniP);
            }
            inkCanvas.Opacity = 1;
            double boundsWidth = GetTouchBoundWidth(e);
            var eraserMultiplier = 1d;
            if (!Settings.Advanced.EraserBindTouchMultiplier && Settings.Advanced.IsSpecialScreen) eraserMultiplier = 1 / Settings.Advanced.TouchMultiplier;
            if (boundsWidth > BoundsWidth)
            {
                isLastTouchEraser = true;
                if (drawingShapeMode == 0 && forceEraser) return;
                if (boundsWidth > BoundsWidth * 2.5)
                {
                    double size = boundsWidth * 3d * GetEraserSizeCoefficient() * eraserMultiplier;
                    inkCanvas.EraserShape = new EllipseStylusShape(size, size);
                    inkCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
                }
                else
                {
                    if (StackPanelPPTControls.Visibility == Visibility.Visible && inkCanvas.Strokes.Count == 0 && Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl)
                    {
                        isLastTouchEraser = false;
                        inkCanvas.EditingMode = InkCanvasEditingMode.GestureOnly;
                        inkCanvas.Opacity = 0.1;
                    }
                    else
                    {
                        inkCanvas.EraserShape = new EllipseStylusShape(5, 5);
                        inkCanvas.EditingMode = InkCanvasEditingMode.EraseByStroke;
                    }
                }
            }
            else
            {
                isLastTouchEraser = false;
                inkCanvas.EraserShape = forcePointEraser ? new EllipseStylusShape(50 * GetEraserSizeCoefficient(), 50 * GetEraserSizeCoefficient()) : new EllipseStylusShape(5, 5);
                if (forceEraser) return;
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            }
        }

        public double GetTouchBoundWidth(TouchEventArgs e)
        {
            var args = e.GetTouchPoint(null).Bounds;
            double value;
            if (!Settings.Advanced.IsQuadIR) value = args.Width;
            else value = Math.Sqrt(args.Width * args.Height); //四边红外
            if (Settings.Advanced.IsSpecialScreen) value *= Settings.Advanced.TouchMultiplier;
            return value;
        }

        //记录触摸设备ID
        private List<int> dec = new List<int>();
        //中心点
        System.Windows.Point centerPoint;
        InkCanvasEditingMode lastInkCanvasEditingMode = InkCanvasEditingMode.Ink;
        bool isSingleFingerDragMode = false;

        //防止衣服误触造成的墨迹消失

        private void inkCanvas_PreviewTouchDown(object sender, TouchEventArgs e)
        {
            dec.Add(e.TouchDevice.Id);
            //设备1个的时候，记录中心点
            if (dec.Count == 1)
            {
                TouchPoint touchPoint = e.GetTouchPoint(inkCanvas);
                centerPoint = touchPoint.Position;

                //记录第一根手指点击时的 StrokeCollection
                lastTouchDownStrokeCollection = inkCanvas.Strokes.Clone();
            }
            //设备两个及两个以上，将画笔功能关闭
            if (dec.Count > 1 || isSingleFingerDragMode || !Settings.Gesture.IsEnableTwoFingerGesture)
            {
                if (isInMultiTouchMode || !Settings.Gesture.IsEnableTwoFingerGesture) return;
                if (inkCanvas.EditingMode != InkCanvasEditingMode.None && inkCanvas.EditingMode != InkCanvasEditingMode.Select)
                {
                    lastInkCanvasEditingMode = inkCanvas.EditingMode;
                    inkCanvas.EditingMode = InkCanvasEditingMode.None;
                }
            }
        }

        private void inkCanvas_PreviewTouchUp(object sender, TouchEventArgs e)
        {
            //手势完成后切回之前的状态
            if (dec.Count > 1)
            {
                if (inkCanvas.EditingMode == InkCanvasEditingMode.None)
                {
                    inkCanvas.EditingMode = lastInkCanvasEditingMode;
                }
            }
            dec.Remove(e.TouchDevice.Id);
            inkCanvas.Opacity = 1;
            if (dec.Count == 0)
            {
                if (lastTouchDownStrokeCollection.Count() != inkCanvas.Strokes.Count() &&
                    !(drawingShapeMode == 9 && !isFirstTouchCuboid))
                {
                    int whiteboardIndex = CurrentWhiteboardIndex;
                    if (currentMode == 0)
                    {
                        whiteboardIndex = 0;
                    }
                    strokeCollections[whiteboardIndex] = lastTouchDownStrokeCollection;
                }
            }
        }
        private void inkCanvas_ManipulationStarting(object sender, ManipulationStartingEventArgs e)
        {
            e.Mode = ManipulationModes.All;
        }

        private void inkCanvas_ManipulationInertiaStarting(object sender, ManipulationInertiaStartingEventArgs e)
        {

        }

        private void Main_Grid_ManipulationCompleted(object sender, ManipulationCompletedEventArgs e)
        {
            if (e.Manipulators.Count() == 0)
            {
                if (forceEraser) return;
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            }
        }

        private void Main_Grid_ManipulationDelta(object sender, ManipulationDeltaEventArgs e)
        {
            if (isInMultiTouchMode || !Settings.Gesture.IsEnableTwoFingerGesture || _lockSmith) return;
            if (dec.Count >= 2 || isSingleFingerDragMode)
            {
                ManipulationDelta md = e.DeltaManipulation;
                Vector trans = md.Translation;  // 获得位移矢量
                double rotate = md.Rotation;  // 获得旋转角度
                Vector scale = md.Scale;  // 获得缩放倍数

                Matrix m = new Matrix();

                // Find center of element and then transform to get current location of center
                FrameworkElement fe = e.Source as FrameworkElement;
                Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
                center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

                // Update matrix to reflect translation/rotation
                if (Settings.Gesture.IsEnableTwoFingerTranslate)
                    m.Translate(trans.X, trans.Y);  // 移动
                if (Settings.Gesture.IsEnableTwoFingerRotation)
                    m.RotateAt(rotate, center.X, center.Y);  // 旋转
                if (Settings.Gesture.IsEnableTwoFingerZoom)
                    m.ScaleAt(scale.X, scale.Y, center.X, center.Y);  // 缩放

                StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
                if (strokes.Count != 0)
                {
                    foreach (Stroke stroke in strokes)
                    {
                        stroke.Transform(m, false);

                        foreach (Circle circle in circles)
                        {
                            if (stroke == circle.Stroke)
                            {
                                circle.R = GetDistance(circle.Stroke.StylusPoints[0].ToPoint(), circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].ToPoint()) / 2;
                                circle.Centroid = new Point((circle.Stroke.StylusPoints[0].X + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].X) / 2,
                                                            (circle.Stroke.StylusPoints[0].Y + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].Y) / 2);
                                break;
                            }
                        }

                        if (Settings.Gesture.IsEnableTwoFingerZoom)
                        {
                            try
                            {
                                stroke.DrawingAttributes.Width *= md.Scale.X;
                                stroke.DrawingAttributes.Height *= md.Scale.Y;
                            }
                            catch { }
                        }
                    }
                }
                else
                {
                    foreach (Stroke stroke in inkCanvas.Strokes)
                    {
                        stroke.Transform(m, false);

                        if (Settings.Gesture.IsEnableTwoFingerZoom)
                        {
                            try
                            {
                                stroke.DrawingAttributes.Width *= md.Scale.X;
                                stroke.DrawingAttributes.Height *= md.Scale.Y;
                            }
                            catch { }
                        }
                    }
                    foreach (Circle circle in circles)
                    {
                        circle.R = GetDistance(circle.Stroke.StylusPoints[0].ToPoint(), circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].ToPoint()) / 2;
                        circle.Centroid = new Point((circle.Stroke.StylusPoints[0].X + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].X) / 2,
                                                    (circle.Stroke.StylusPoints[0].Y + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].Y) / 2);
                    }
                }
            }
        }

        #endregion Touch Events

        #region PowerPoint

        public static Microsoft.Office.Interop.PowerPoint.Application pptApplication = null;
        public static Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
        public static Microsoft.Office.Interop.PowerPoint.Slides slides = null;
        public static Microsoft.Office.Interop.PowerPoint.Slide slide = null;
        public static int slidescount = 0;
        private void BtnCheckPPT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pptApplication = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                //pptApplication.SlideShowWindows[1].View.Next();
                if (pptApplication != null)
                {
                    //获得演示文稿对象
                    presentation = pptApplication.ActivePresentation;
                    pptApplication.SlideShowBegin += PptApplication_SlideShowBegin;
                    pptApplication.SlideShowNextSlide += PptApplication_SlideShowNextSlide;
                    pptApplication.SlideShowEnd += PptApplication_SlideShowEnd;
                    // 获得幻灯片对象集合
                    slides = presentation.Slides;
                    // 获得幻灯片的数量
                    slidescount = slides.Count;
                    memoryStreams = new MemoryStream[slidescount + 2];
                    // 获得当前选中的幻灯片
                    try
                    {
                        // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                        // 然而在阅读模式下，这种方式会出现异常
                        slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                    }
                    catch
                    {
                        // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                        slide = pptApplication.SlideShowWindows[1].View.Slide;
                    }
                }

                if (pptApplication == null) throw new Exception();
                //BtnCheckPPT.Visibility = Visibility.Collapsed;
                StackPanelPPTControls.Visibility = Visibility.Visible;
            }
            catch
            {
                //BtnCheckPPT.Visibility = Visibility.Visible;
                StackPanelPPTControls.Visibility = Visibility.Collapsed;
                MessageBox.Show("未找到幻灯片");
            }
        }
        private void ToggleSwitchSupportWPS_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.IsSupportWPS = ToggleSwitchSupportWPS.IsOn;
            SaveSettingsToFile();
        }

        public static bool isWPSSupportOn => Settings.PowerPointSettings.IsSupportWPS;

        public static bool IsShowingRestoreHiddenSlidesWindow = false;

        public static bool IsNotifyPreviousPageWindowShown = false;

        private void TimerCheckPPT_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (IsShowingRestoreHiddenSlidesWindow) return;
            try
            {
                Process[] processes = Process.GetProcessesByName("wpp");
                if (processes.Length > 0 && !isWPSSupportOn)
                {
                    return;
                }

                //使用下方提前创建 PowerPoint 实例，将导致 PowerPoint 不再有启动界面
                //pptApplication = (Microsoft.Office.Interop.PowerPoint.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("91493441-5A91-11CF-8700-00AA0060263B")));
                //new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowBegin").AddEventHandler(pptApplication, new EApplication_SlideShowBeginEventHandler(this.PptApplication_SlideShowBegin));
                //new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowEnd").AddEventHandler(pptApplication, new EApplication_SlideShowEndEventHandler(this.PptApplication_SlideShowEnd));
                //new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowNextSlide").AddEventHandler(pptApplication, new EApplication_SlideShowNextSlideEventHandler(this.PptApplication_SlideShowNextSlide));
                //ConfigHelper.Instance.IsInitApplicationSuccessful = true;

                pptApplication = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");

                if (pptApplication != null)
                {
                    timerCheckPPT.Stop();
                    //获得演示文稿对象
                    presentation = pptApplication.ActivePresentation;
                    pptApplication.PresentationClose += PptApplication_PresentationClose;
                    pptApplication.SlideShowBegin += PptApplication_SlideShowBegin;
                    pptApplication.SlideShowNextSlide += PptApplication_SlideShowNextSlide;
                    pptApplication.SlideShowEnd += PptApplication_SlideShowEnd;
                    // 获得幻灯片对象集合
                    slides = presentation.Slides;

                    // 获得幻灯片的数量
                    slidescount = slides.Count;
                    memoryStreams = new MemoryStream[slidescount + 2];
                    // 获得当前选中的幻灯片
                    try
                    {
                        // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                        // 然而在阅读模式下，这种方式会出现异常
                        slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                    }
                    catch
                    {
                        // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                        slide = pptApplication.SlideShowWindows[1].View.Slide;
                    }
                }

                if (pptApplication == null) return;
                //BtnCheckPPT.Visibility = Visibility.Collapsed;

                // 跳转到上次播放页
                if (Settings.PowerPointSettings.IsNotifyPreviousPage)
                    Application.Current.Dispatcher.BeginInvoke(() =>
                        {
                            string defaultFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                                                       @"\Ink Canvas Strokes\Auto Saved\Presentations\";
                            string folderPath = defaultFolderPath + presentation.Name + "_" + presentation.Slides.Count;
                            if (File.Exists(folderPath + "/Position") & !IsNotifyPreviousPageWindowShown) //判断是否已存在NotifyPreviousPage窗口
                            {
                                if (int.TryParse(File.ReadAllText(folderPath + "/Position"), out var page))
                                {
                                    IsNotifyPreviousPageWindowShown = true;
                                    if (page <= 0) return;
                                    new YesOrNoNotificationWindow($"上次播放到了第 {page} 页, 是否立即跳转", () =>
                                    {
                                        if (pptApplication.SlideShowWindows.Count >= 1)
                                        {
                                            // 如果已经播放了的话, 跳转
                                            presentation.SlideShowWindow.View.GotoSlide(page);
                                        }
                                        else
                                        {
                                            presentation.Windows[1].View.GotoSlide(page);
                                        }
                                    }).ShowDialog();
                                }
                            }
                        }, DispatcherPriority.Normal);


                //检查是否有隐藏幻灯片
                if (Settings.PowerPointSettings.IsNotifyHiddenPage)
                {
                    bool isHaveHiddenSlide = false;
                    foreach (Slide slide in slides)
                    {
                        if (slide.SlideShowTransition.Hidden == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            isHaveHiddenSlide = true;
                            break;
                        }
                    }

                    Application.Current.Dispatcher.BeginInvoke(() =>
                    {
                        if (isHaveHiddenSlide && !IsShowingRestoreHiddenSlidesWindow)
                        {
                            IsShowingRestoreHiddenSlidesWindow = true;
                            new YesOrNoNotificationWindow("检测到此演示文稿包含隐藏的幻灯片，是否取消隐藏？",
                                () =>
                                {
                                    foreach (Slide slide in slides)
                                    {
                                        if (slide.SlideShowTransition.Hidden ==
                                            Microsoft.Office.Core.MsoTriState.msoTrue)
                                        {
                                            slide.SlideShowTransition.Hidden =
                                                Microsoft.Office.Core.MsoTriState.msoFalse;
                                        }
                                    }
                                }).ShowDialog();
                        }



                        BtnPPTSlideShow.Visibility = Visibility.Visible;
                    }, DispatcherPriority.Normal);
                }

                //如果检测到已经开始放映，则立即进入画板模式
                if (pptApplication.SlideShowWindows.Count >= 1)
                {
                    PptApplication_SlideShowBegin(pptApplication.SlideShowWindows[1]);
                }
            }
            catch
            {
                //StackPanelPPTControls.Visibility = Visibility.Collapsed;
                Application.Current.Dispatcher.Invoke(() =>
                {
                    BtnPPTSlideShow.Visibility = Visibility.Collapsed;
                });
                timerCheckPPT.Start();
            }
        }

        private void PptApplication_PresentationClose(Presentation Pres)
        {
            pptApplication.PresentationClose -= PptApplication_PresentationClose;
            pptApplication.SlideShowBegin -= PptApplication_SlideShowBegin;
            pptApplication.SlideShowNextSlide -= PptApplication_SlideShowNextSlide;
            pptApplication.SlideShowEnd -= PptApplication_SlideShowEnd;
            pptApplication = null;
            timerCheckPPT.Start();
            if (BtnPPTSlideShowEnd.Visibility != Visibility.Collapsed)
            {
                PptApplication_SlideShowEnd(Pres);
            }
        }

        bool isPresentationHaveBlackSpace = false;


        private string pptName = null;
        int currentShowPosition = -1;
        //bool isButtonBackgroundTransparent = true; //此变量仅用于保存用于幻灯片放映时的优化
        private void PptApplication_SlideShowBegin(SlideShowWindow Wn)
        {
            LogHelper.WriteLogToFile("PowerPoint Application Slide Show Begin", LogHelper.LogType.Event);
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (currentMode == 1)
                {
                    // 退出画板模式
                    BtnSwitch_Click(null, null);
                }

                //调整颜色
                double screenRatio = SystemParameters.PrimaryScreenWidth / SystemParameters.PrimaryScreenHeight;
                if (Math.Abs(screenRatio - 16.0 / 9) <= -0.01)
                {
                    if (Wn.Presentation.PageSetup.SlideWidth / Wn.Presentation.PageSetup.SlideHeight < 1.65)
                    {
                        isPresentationHaveBlackSpace = true;
                        //isButtonBackgroundTransparent = ToggleSwitchTransparentButtonBackground.IsOn;

                        if (BtnSwitchTheme.Content.ToString() == "深色")
                        {
                            //Light
                            BtnExit.Foreground = Brushes.White;
                            SymbolIconBtnColorBlackContent.Foreground = Brushes.White;
                            ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
                            //BtnExit.Background = new SolidColorBrush(StringToColor("#AACCCCCC"));
                        }
                        else
                        {
                            //Dark
                            //BtnExit.Background = new SolidColorBrush(StringToColor("#AA555555"));
                        }
                    }
                }
                else if (screenRatio == -256 / 135)
                {

                }

                slidescount = Wn.Presentation.Slides.Count;
                previousSlideID = 0;
                memoryStreams = new MemoryStream[slidescount + 2];

                pptName = Wn.Presentation.Name;
                LogHelper.NewLog("Name: " + Wn.Presentation.Name);
                LogHelper.NewLog("Slides Count: " + slidescount.ToString());

                //检查是否有已有墨迹，并加载
                if (Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint)
                {
                    string defaultFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\Auto Saved\Presentations\";
                    if (Directory.Exists(defaultFolderPath + Wn.Presentation.Name + "_" + Wn.Presentation.Slides.Count))
                    {
                        LogHelper.WriteLogToFile("Found saved strokes", LogHelper.LogType.Trace);
                        FileInfo[] files = new DirectoryInfo(defaultFolderPath + Wn.Presentation.Name + "_" + Wn.Presentation.Slides.Count).GetFiles();
                        int count = 0;
                        foreach (FileInfo file in files)
                        {
                            if (file.Name != "Position")
                            {
                                int i = -1;
                                try
                                {
                                    i = int.Parse(System.IO.Path.GetFileNameWithoutExtension(file.Name));
                                    //var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);
                                    //MemoryStream ms = new MemoryStream(File.ReadAllBytes(file.FullName));
                                    //new StrokeCollection(fs).Save(ms);
                                    //ms.Position = 0;
                                    memoryStreams[i] = new MemoryStream(File.ReadAllBytes(file.FullName));
                                    memoryStreams[i].Position = 0;
                                    count++;
                                }
                                catch (Exception ex)
                                {
                                    LogHelper.WriteLogToFile(string.Format("Failed to load strokes on Slide {0}\n{1}", i, ex.ToString()), LogHelper.LogType.Error);
                                }
                            }
                        }
                        LogHelper.WriteLogToFile(string.Format("Loaded {0} saved strokes", count.ToString()));
                    }
                }

                pointDesktop = new Point(ViewboxFloatingBar.Margin.Left, ViewboxFloatingBar.Margin.Top);
                pointPPT = new Point(-1, -1);

                StackPanelPPTControls.Visibility = Visibility.Visible;
                BtnPPTSlideShow.Visibility = Visibility.Collapsed;
                BtnPPTSlideShowEnd.Visibility = Visibility.Visible;
                ViewBoxStackPanelMain.Margin = new Thickness(10, 10, 10, 10);
                lockSmithDesktop = _lockSmith;
                ChangeLockSmithState(!Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode);

                if (Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow && Main_Grid.Background == Brushes.Transparent)
                {
                    if (currentMode != 0)
                    {
                        currentMode = 0;
                        GridBackgroundCover.Visibility = Visibility.Collapsed;

                        //SaveStrokes();
                        ClearStrokes(true);

                        if (BtnSwitchTheme.Content.ToString() == "浅色")
                        {
                            BtnSwitch.Content = "黑板";
                        }
                        else
                        {
                            BtnSwitch.Content = "白板";
                        }
                        StackPanelPPTButtons.Visibility = Visibility.Visible;
                    }
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                }
                //if (GridBackgroundCover.Visibility == Visibility.Visible)
                //{
                //    SaveStrokes();
                //    currentMode = 0;
                //    GridBackgroundCover.Visibility = Visibility.Hidden;
                //}

                ClearStrokes(true);

                SetBorderFloatingBarMainControlsVisibility(true, false);
                BorderPenColorRed_MouseUp(BorderPenColorRed, null);

                if (Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow == false)
                {
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                }

                isEnteredSlideShowEndEvent = false;
                PptNavigationTextBlock.Text = $"{Wn.View.CurrentShowPosition}/{Wn.Presentation.Slides.Count}";
                LogHelper.NewLog("PowerPoint Slide Show Loading process complete");

                new Thread(new ThreadStart(() =>
                {
                    Thread.Sleep(100);
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth * FloatingBarScale) / 2, SystemParameters.PrimaryScreenHeight - 60 + ViewboxFloatingBar.ActualHeight * (1 - FloatingBarScale), -2000, -200);
                    });
                })).Start();
            });

            Application.Current.Dispatcher.Invoke(() =>
            {
                UpdateWindowTitle();
            });
            //previousSlideID = Wn.View.CurrentShowPosition;
            ////检查是否有已有墨迹，并加载当前页
            //if (Settings.Automation.IsAutoSaveStrokesInPowerPoint)
            //{
            //    try
            //    {
            //        if (memoryStreams[Wn.View.CurrentShowPosition].Length > 0)
            //        {
            //            Application.Current.Dispatcher.Invoke(() =>
            //            {
            //                inkCanvas.Strokes = new System.Windows.Ink.StrokeCollection(memoryStreams[Wn.View.CurrentShowPosition]);
            //            });
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        LogHelper.WriteLogToFile(string.Format("Failed to load strokes for current slide (Slide {0})\n{1}", Wn.View.CurrentShowPosition, ex.ToString()), LogHelper.LogType.Error);
            //    }
            //}
        }

        bool isEnteredSlideShowEndEvent = false; //防止重复调用本函数导致墨迹保存失效
        private void PptApplication_SlideShowEnd(Presentation Pres)
        {
            IsNotifyPreviousPageWindowShown = false;
            LogHelper.WriteLogToFile(string.Format("PowerPoint Slide Show End"), LogHelper.LogType.Event);
            if (isEnteredSlideShowEndEvent)
            {
                LogHelper.WriteLogToFile("Detected previous entrance, returning");
                return;
            }
            isEnteredSlideShowEndEvent = true;
            if (Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint)
            {
                string defaultFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\Auto Saved\Presentations\";
                string folderPath = defaultFolderPath + Pres.Name + "_" + Pres.Slides.Count;
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                try
                {
                    File.WriteAllText(folderPath + "/Position", previousSlideID.ToString());
                }
                catch { }
                Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        MemoryStream ms = new MemoryStream();
                        inkCanvas.Strokes.Save(ms);
                        ms.Position = 0;
                        memoryStreams[currentShowPosition] = ms;
                    }
                    catch { }
                });
                for (int i = 1; i <= Pres.Slides.Count; i++)
                {
                    if (memoryStreams[i] != null)
                    {
                        try
                        {
                            if (memoryStreams[i].Length > 8)
                            {
                                byte[] srcBuf = new Byte[memoryStreams[i].Length];
                                //MessageBox.Show(memoryStreams[i].Length.ToString());
                                int byteLength = memoryStreams[i].Read(srcBuf, 0, srcBuf.Length);
                                File.WriteAllBytes(folderPath + @"\" + i.ToString("0000") + ".icstk", srcBuf);
                                LogHelper.WriteLogToFile(string.Format("Saved strokes for Slide {0}, size={1}, byteLength={2}", i.ToString(), memoryStreams[i].Length, byteLength));
                            }
                            else
                            {
                                File.Delete(folderPath + @"\" + i.ToString("0000") + ".icstk");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogHelper.WriteLogToFile(string.Format("Failed to save strokes for Slide {0}\n{1}", i, ex.ToString()), LogHelper.LogType.Error);
                            File.Delete(folderPath + @"\" + i.ToString("0000") + ".icstk");
                        }
                    }
                }
            }

            Application.Current.Dispatcher.Invoke(() =>
            {
                isPresentationHaveBlackSpace = false;

                if (BtnSwitchTheme.Content.ToString() == "深色")
                {
                    //Light
                    BtnExit.Foreground = Brushes.Black;
                    SymbolIconBtnColorBlackContent.Foreground = Brushes.White;
                    ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
                }
                else
                {
                    //Dark
                }

                BtnPPTSlideShow.Visibility = Visibility.Visible;
                BtnPPTSlideShowEnd.Visibility = Visibility.Collapsed;
                StackPanelPPTControls.Visibility = Visibility.Collapsed;
                ViewBoxStackPanelMain.Margin = new Thickness(10, 10, 10, 55);
                ChangeLockSmithState(lockSmithDesktop);

                if (currentMode != 0)
                {
                    currentMode = 0;
                    GridBackgroundCover.Visibility = Visibility.Collapsed;

                    //SaveStrokes();
                    ClearStrokes(true);
                    //RestoreStrokes(true);

                    if (BtnSwitchTheme.Content.ToString() == "浅色")
                    {
                        BtnSwitch.Content = "黑板";
                    }
                    else
                    {
                        BtnSwitch.Content = "白板";
                    }
                    StackPanelPPTButtons.Visibility = Visibility.Visible;
                }
                //if (GridBackgroundCover.Visibility == Visibility.Visible)
                //{
                //    SaveStrokes();
                //}


                ClearStrokes(true);

                if (Main_Grid.Background != Brushes.Transparent)
                {
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                }

                if (pointDesktop != new Point(-1, -1))
                {
                    ViewboxFloatingBar.Margin = new Thickness(pointDesktop.X, pointDesktop.Y, -2000, -200);
                    if (Settings.Appearance.IsAutoCollapseFloatBar)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            SetBorderFloatingBarMainControlsVisibility(false);
                        }), DispatcherPriority.Loaded);
                    }

                }
            });

            Application.Current.Dispatcher.Invoke(() =>
            {
                UpdateWindowTitle();
            });
        }

        int previousSlideID = 0;
        MemoryStream[] memoryStreams = new MemoryStream[50];

        private void PptApplication_SlideShowNextSlide(SlideShowWindow Wn)
        {
            LogHelper.WriteLogToFile(string.Format("PowerPoint Next Slide (Slide {0})", Wn.View.CurrentShowPosition), LogHelper.LogType.Event);
            if (Wn.View.CurrentShowPosition != previousSlideID)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    MemoryStream ms = new MemoryStream();
                    inkCanvas.Strokes.Save(ms);
                    ms.Position = 0;
                    memoryStreams[previousSlideID] = ms;

                    if (inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber && Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint && !_isPptClickingBtnTurned)
                        SaveScreenShot(true, Wn.Presentation.Name + "/" + Wn.View.CurrentShowPosition);
                    _isPptClickingBtnTurned = false;

                    ClearStrokes(true);
                    timeMachine.ClearStrokeHistory();

                    try
                    {
                        if (memoryStreams[Wn.View.CurrentShowPosition] != null && memoryStreams[Wn.View.CurrentShowPosition].Length > 0)
                        {
                            inkCanvas.Strokes.Add(new StrokeCollection(memoryStreams[Wn.View.CurrentShowPosition]));
                        }
                        currentShowPosition = Wn.View.CurrentShowPosition;
                    }
                    catch
                    { }

                    PptNavigationTextBlock.Text = $"{Wn.View.CurrentShowPosition}/{Wn.Presentation.Slides.Count}";
                });
                previousSlideID = Wn.View.CurrentShowPosition;

            }
        }

        private bool _isPptClickingBtnTurned = false;

        private void BtnPPTSlidesUp_Click(object sender, RoutedEventArgs e)
        {
            if (currentMode == 1)
            {
                GridBackgroundCover.Visibility = Visibility.Collapsed;
                currentMode = 0;
            }

            _isPptClickingBtnTurned = true;
            if (inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber &&
                Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint)
                SaveScreenShot(true, pptApplication.SlideShowWindows[1].Presentation.Name + "/" + pptApplication.SlideShowWindows[1].View.CurrentShowPosition);
            try
            {
                new Thread(new ThreadStart(() =>
                {
                    pptApplication.SlideShowWindows[1].Activate();
                    pptApplication.SlideShowWindows[1].View.Previous();
                })).Start();
            }
            catch
            {
                //BtnCheckPPT.Visibility = Visibility.Visible;
                StackPanelPPTControls.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnPPTSlidesDown_Click(object sender, RoutedEventArgs e)
        {
            if (currentMode == 1)
            {
                GridBackgroundCover.Visibility = Visibility.Collapsed;
                currentMode = 0;
            }
            _isPptClickingBtnTurned = true;
            if (inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber &&
                Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint)
                SaveScreenShot(true, pptApplication.SlideShowWindows[1].Presentation.Name + "/" + pptApplication.SlideShowWindows[1].View.CurrentShowPosition);
            try
            {
                new Thread(new ThreadStart(() =>
                {
                    pptApplication.SlideShowWindows[1].Activate();
                    pptApplication.SlideShowWindows[1].View.Next();
                })).Start();
            }
            catch
            {
                StackPanelPPTControls.Visibility = Visibility.Collapsed;
            }
        }


        private async void PPTNavigationBtn_Click(object sender, MouseButtonEventArgs e)
        {
            Main_Grid.Background = new SolidColorBrush(StringToColor("#01FFFFFF"));
            BtnHideInkCanvas_Click(sender, e);
            try
            {
                pptApplication.SlideShowWindows[1].SlideNavigation.Visible = true;
            }
            catch
            {
                ShowNotification("此功能不支持阅读模式，请在幻灯片放映模式下使用。");
            }
            // 控制居中
            if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
            {
                if (ViewboxFloatingBar.Margin == new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth * FloatingBarScale) / 2, SystemParameters.PrimaryScreenHeight - 60 + ViewboxFloatingBar.ActualHeight * (1 - FloatingBarScale), -2000, -200))
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth * FloatingBarScale) / 2, SystemParameters.PrimaryScreenHeight - 60 + ViewboxFloatingBar.ActualHeight * (1 - FloatingBarScale), -2000, -200);
                    }), DispatcherPriority.Render);
                }
            }
        }

        private void BtnPPTSlideShow_Click(object sender, RoutedEventArgs e)
        {
            new Thread(new ThreadStart(() =>
            {
                try
                {
                    presentation.SlideShowSettings.Run();
                }
                catch { }
            })).Start();
        }

        private void BtnPPTSlideShowEnd_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                try
                {
                    MemoryStream ms = new MemoryStream();
                    inkCanvas.Strokes.Save(ms);
                    ms.Position = 0;
                    memoryStreams[pptApplication.SlideShowWindows[1].View.CurrentShowPosition] = ms;
                    timeMachine.ClearStrokeHistory();
                    IsNotifyPreviousPageWindowShown = false;
                }
                catch { }
            });
            new Thread(new ThreadStart(() =>
            {
                try
                {
                    pptApplication.SlideShowWindows[1].View.Exit();
                }
                catch { }
            })).Start();
        }

        #endregion

        #region Settings

        #region Behavior

        private void ToggleSwitchRunAtStartup_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            if (ToggleSwitchRunAtStartup.IsOn)
            {
                StartAutomaticallyCreate("InkCanvas");
            }
            else
            {
                StartAutomaticallyDel("InkCanvas");
            }
        }

        private void ToggleSwitchSupportPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.PowerPointSupport = ToggleSwitchSupportPowerPoint.IsOn;
            SaveSettingsToFile();

            if (Settings.PowerPointSettings.PowerPointSupport)
            {
                timerCheckPPT.Start();
            }
            else
            {
                timerCheckPPT.Stop();
            }
        }

        private void ToggleSwitchShowCanvasAtNewSlideShow_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow = ToggleSwitchShowCanvasAtNewSlideShow.IsOn;
            SaveSettingsToFile();
        }

        #endregion

        #region Startup

        private void ToggleSwitchAutoHideCanvas_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Startup.IsAutoHideCanvas = ToggleSwitchAutoHideCanvas.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchAutoEnterModeFinger_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Startup.IsAutoEnterModeFinger = ToggleSwitchAutoEnterModeFinger.IsOn;
            SaveSettingsToFile();
        }

        #endregion

        #region Appearance


        private void SideControlOpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void ToggleSwitchAutoCollapseFloatBar_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Appearance.IsAutoCollapseFloatBar = ToggleSwitchAutoCollapseFloatBar.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchFloatBarShowOnRight_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Appearance.IsFloatBarShowOnRight = ToggleSwitchFloatBarShowOnRight.IsOn;
            SaveSettingsToFile();
            UpdateFloatBarExpansionDirection();
        }

        private void ToggleSwitchRememberFloatBarPosition_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Appearance.IsRememberFloatBarPosition = ToggleSwitchRememberFloatBarPosition.IsOn;
            SaveSettingsToFile();
            if (Settings.Appearance.IsRememberFloatBarPosition)
            {
                SaveFloatBarPositionToFile();
            }
        }

        private void UpdateFloatBarExpansionDirection()
        {
            if (!isLoaded) return;
            if (StackPanelFloatBarContainer == null) return;

            var isFloatBarShowOnRight = Settings.Appearance.IsFloatBarShowOnRight;

            // 通过子元素顺序控制展开方向，避免 FlowDirection 反转导致方向判断异常
            StackPanelFloatBarContainer.FlowDirection = FlowDirection.LeftToRight;
            BorderFloatingBarMainControls.FlowDirection = FlowDirection.LeftToRight;
            BorderFloatingBarMainControls.RenderTransformOrigin = isFloatBarShowOnRight ? new Point(1, 0.5) : new Point(0, 0.5);
            BorderFloatingBarMainControls.Margin = isFloatBarShowOnRight ? new Thickness(0, 0, 5, 0) : new Thickness(5, 0, 0, 0);

            StackPanelFloatBarContainer.Children.Remove(BorderFloatingBarEmoji);
            StackPanelFloatBarContainer.Children.Remove(BorderFloatingBarMainControls);

            if (isFloatBarShowOnRight)
            {
                StackPanelFloatBarContainer.Children.Add(BorderFloatingBarMainControls);
                StackPanelFloatBarContainer.Children.Add(BorderFloatingBarEmoji);
            }
            else
            {
                StackPanelFloatBarContainer.Children.Add(BorderFloatingBarEmoji);
                StackPanelFloatBarContainer.Children.Add(BorderFloatingBarMainControls);
            }

            // 重置位置
            if (currentMode != 0 || BtnPPTSlideShowEnd.Visibility == Visibility.Visible) return;
            if (isFloatBarShowOnRight)
            {
                ViewboxFloatingBar.Margin = new Thickness(SystemParameters.WorkArea.Left + SystemParameters.WorkArea.Width - 80 - ViewboxFloatingBar.ActualWidth * FloatingBarScale, SystemParameters.WorkArea.Top + SystemParameters.WorkArea.Height - 80 + ViewboxFloatingBar.ActualHeight * (1 - FloatingBarScale), -2000, -200);
            }
            else
            {
                ViewboxFloatingBar.Margin = new Thickness(SystemParameters.WorkArea.Left + 80, SystemParameters.WorkArea.Top + SystemParameters.WorkArea.Height - 80 + ViewboxFloatingBar.ActualHeight * (1 - FloatingBarScale), -2000, -200);
            }
        }

        private void ToggleSwitchShowButtonExit_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Appearance.IsShowExitButton = ToggleSwitchShowButtonExit.IsOn;
            SaveSettingsToFile();

            if (ToggleSwitchShowButtonExit.IsOn)
            {
                BtnExit.Visibility = Visibility.Visible;
            }
            else
            {
                BtnExit.Visibility = Visibility.Collapsed;
            }
        }

        private void ToggleSwitchShowButtonEraser_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Appearance.IsShowEraserButton = ToggleSwitchShowButtonEraser.IsOn;
            SaveSettingsToFile();

            if (ToggleSwitchShowButtonEraser.IsOn)
            {
                BtnErase.Visibility = Visibility.Visible;
            }
            else
            {
                BtnErase.Visibility = Visibility.Collapsed;
            }
        }
        private void ToggleSwitchShowButtonPPTNavigation_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.IsShowPPTNavigation = ToggleSwitchShowButtonPPTNavigation.IsOn;
            SaveSettingsToFile();

            ViewboxPPTSidesControl.Visibility =
                Settings.PowerPointSettings.IsShowPPTNavigation ? Visibility.Visible : Visibility.Collapsed;
            ViewboxPPTRightBottom.Visibility =
                Settings.PowerPointSettings.IsShowPPTNavigation ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ToggleSwitchShowVerticalPPTNavigation_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.IsShowVerticalPPTNavigation = ToggleSwitchShowVerticalPPTNavigation.IsOn;
            SaveSettingsToFile();

            ViewboxPPTLeftCenter.Visibility =
                Settings.PowerPointSettings.IsShowVerticalPPTNavigation ? Visibility.Visible : Visibility.Collapsed;
            ViewboxPPTRightCenter.Visibility =
                Settings.PowerPointSettings.IsShowVerticalPPTNavigation ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ComboBoxTheme_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Appearance.Theme = ComboBoxTheme.SelectedIndex;
            SystemEvents_UserPreferenceChanged(null, null);
            SaveSettingsToFile();
        }


        private void ToggleSwitchShowButtonHideControl_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Appearance.IsShowHideControlButton = ToggleSwitchShowButtonHideControl.IsOn;
            SaveSettingsToFile();

            if (ToggleSwitchShowButtonHideControl.IsOn)
            {
                BtnHideControl.Visibility = Visibility.Visible;
            }
            else
            {
                BtnHideControl.Visibility = Visibility.Collapsed;
            }
        }

        private void ToggleSwitchShowButtonLRSwitch_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Appearance.IsShowLRSwitchButton = ToggleSwitchShowButtonLRSwitch.IsOn;
            SaveSettingsToFile();

            if (ToggleSwitchShowButtonLRSwitch.IsOn)
            {
                BtnSwitchSide.Visibility = Visibility.Visible;
            }
            else
            {
                BtnSwitchSide.Visibility = Visibility.Collapsed;
            }
        }

        private void ToggleSwitchShowButtonModeFinger_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Appearance.IsShowModeFingerToggleSwitch = ToggleSwitchShowButtonModeFinger.IsOn;
            SaveSettingsToFile();

            if (ToggleSwitchShowButtonModeFinger.IsOn)
            {
                StackPanelModeFinger.Visibility = Visibility.Visible;
            }
            else
            {
                StackPanelModeFinger.Visibility = Visibility.Collapsed;
            }
        }

        private void ToggleSwitchTransparentButtonBackground_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Appearance.IsTransparentButtonBackground = ToggleSwitchTransparentButtonBackground.IsOn;
            if (Settings.Appearance.IsTransparentButtonBackground)
            {
                BtnExit.Background = new SolidColorBrush(StringToColor("#7F909090"));
            }
            else
            {
                if (BtnSwitchTheme.Content.ToString() == "深色")
                {
                    //Light
                    BtnExit.Background = new SolidColorBrush(StringToColor("#FFCCCCCC"));
                }
                else
                {
                    //Dark
                    BtnExit.Background = new SolidColorBrush(StringToColor("#FF555555"));
                }
            }

            SaveSettingsToFile();
        }

        private void ToggleSwitchShowCursor_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Canvas.IsShowCursor = ToggleSwitchShowCursor.IsOn;
            inkCanvas_EditingModeChanged(inkCanvas, null);

            SaveSettingsToFile();
        }

        #endregion

        #region Canvas

        private void ComboBoxPenStyle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Canvas.InkStyle = ComboBoxPenStyle.SelectedIndex;
            SaveSettingsToFile();
        }

        private void ComboBoxEraserSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Canvas.EraserSize = ComboBoxEraserSize.SelectedIndex - 2;
            SaveSettingsToFile();
        }


        private void ComboBoxEraserType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Canvas.EraserType = ComboBoxEraserType.SelectedIndex;
            SaveSettingsToFile();
        }

        private void FloatingBarScaleSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var slider = (Slider)sender;
            double val = slider.Value / 100.0;
            double clampedVal = val < 0.5 ? 0.5 : val > 1.0 ? 1.0 : val;
            if (slider.Value != clampedVal * 100)
            {
                slider.Value = clampedVal * 100;
                return;
            }
            if (!isLoaded) return;
            Settings.Appearance.ViewboxFloatingBarScaleTransformValue = clampedVal;
            SaveSettingsToFile();

            ViewboxFloatingBarScaleTransform.ScaleX = clampedVal;
            ViewboxFloatingBarScaleTransform.ScaleY = clampedVal;
        }

        private void InkWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;

            drawingAttributes.Height = ((Slider)sender).Value / 2;
            drawingAttributes.Width = ((Slider)sender).Value / 2;

            Settings.Canvas.InkWidth = ((Slider)sender).Value / 2;

            SaveSettingsToFile();
        }

        private void ComboBoxHyperbolaAsymptoteOption_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Canvas.HyperbolaAsymptoteOption = (OptionalOperation)ComboBoxHyperbolaAsymptoteOption.SelectedIndex;
            SaveSettingsToFile();
        }

        #endregion

        #region Automation

        private void ToggleSwitchAutoKillPptService_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.IsAutoKillPptService = ToggleSwitchAutoKillPptService.IsOn;
            SaveSettingsToFile();

            if (Settings.Automation.IsAutoKillEasiNote || Settings.Automation.IsAutoKillPptService)
            {
                timerKillProcess.Start();
            }
            else
            {
                timerKillProcess.Stop();
            }
        }

        private void ToggleSwitchAutoKillEasiNote_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.IsAutoKillEasiNote = ToggleSwitchAutoKillEasiNote.IsOn;
            SaveSettingsToFile();

            if (Settings.Automation.IsAutoKillEasiNote || Settings.Automation.IsAutoKillPptService)
            {
                timerKillProcess.Start();
            }
            else
            {
                timerKillProcess.Stop();
            }
        }

        private void ToggleSwitchSaveScreenshotsInDateFolders_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.IsSaveScreenshotsInDateFolders = ToggleSwitchSaveScreenshotsInDateFolders.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesAtScreenshot_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.IsAutoSaveStrokesAtScreenshot = ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn;
            ToggleSwitchAutoSaveStrokesAtClear.Header =
                ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn ? "清屏时自动截图并保存墨迹" : "清屏时自动截图";
            SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesAtClear_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.IsAutoSaveStrokesAtClear = ToggleSwitchAutoSaveStrokesAtClear.IsOn;
            SaveSettingsToFile();
        }


        private void ToggleSwitchExitingWritingMode_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.IsAutoClearWhenExitingWritingMode = ToggleSwitchClearExitingWritingMode.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchHideStrokeWhenSelecting_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Canvas.HideStrokeWhenSelecting = ToggleSwitchHideStrokeWhenSelecting.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchUsingWhiteboard_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Canvas.UsingWhiteboard = ToggleSwitchUsingWhiteboard.IsOn;
            if (!Settings.Canvas.UsingWhiteboard)
            {
                BtnSwitchTheme.Content = "浅色";
            }
            else
            {
                BtnSwitchTheme.Content = "深色";
            }
            BtnSwitchTheme_Click(sender, e);
            SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint = ToggleSwitchAutoSaveStrokesInPowerPoint.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyPreviousPage_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.PowerPointSettings.IsNotifyPreviousPage = ToggleSwitchNotifyPreviousPage.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyHiddenPage_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.PowerPointSettings.IsNotifyHiddenPage = ToggleSwitchNotifyHiddenPage.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchNoStrokeClearInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.PowerPointSettings.IsNoClearStrokeOnSelectWhenInPowerPoint = ToggleSwitchNoStrokeClearInPowerPoint.IsOn;
            SaveSettingsToFile();
        }


        private void ToggleSwitchShowStrokeOnSelectInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.PowerPointSettings.IsShowStrokeOnSelectInPowerPoint = ToggleSwitchShowStrokeOnSelectInPowerPoint.IsOn;
            SaveSettingsToFile();
        }

        private void SideControlMinimumAutomationSlider_ValueChanged(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Automation.MinimumAutomationStrokeNumber = (int)SideControlMinimumAutomationSlider.Value;
            SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveScreenShotInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint = ToggleSwitchAutoSaveScreenShotInPowerPoint.IsOn;
            SaveSettingsToFile();
        }

        #endregion

        #region Gesture


        private void ToggleSwitchEnableFingerGestureSlideShowControl_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl = ToggleSwitchEnableFingerGestureSlideShowControl.IsOn;

            SaveSettingsToFile();
        }

        private void ToggleSwitchDisableLockSmithByDefault_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Gesture.IsDisableLockSmithByDefault = ToggleSwitchDisableLockSmithByDefault.IsOn;

            SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerZoom_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Gesture.IsEnableTwoFingerZoom = ToggleSwitchEnableTwoFingerZoom.IsOn;

            SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerTranslate_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Gesture.IsEnableTwoFingerTranslate = ToggleSwitchEnableTwoFingerTranslate.IsOn;

            SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerRotation_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.Gesture.IsEnableTwoFingerRotation = ToggleSwitchEnableTwoFingerRotation.IsOn;
            Settings.Gesture.IsEnableTwoFingerRotationOnSelection = ToggleSwitchEnableTwoFingerRotationOnSelection.IsOn;

            SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerGestureInPresentationMode_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;

            Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode = ToggleSwitchEnableTwoFingerGestureInPresentationMode.IsOn;

            SaveSettingsToFile();
        }

        #endregion

        #region Reset

        private void BtnResetToDefault_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FlyoutResetToDefault.Hide();
                isLoaded = false;
                File.Delete("settings.json");
                Settings = new Settings();
                LoadSettings(false);
                isLoaded = true;
            }
            catch { }
        }

        #endregion

        #region Ink To Shape

        private void ToggleSwitchEnableInkToShape_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.InkToShape.IsInkToShapeEnabled = ToggleSwitchEnableInkToShape.IsOn;
            SaveSettingsToFile();
        }

        private void LineNormalizationThresholdSlider_ValueChanged(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.InkToShape.LineNormalizationThreshold = (double)LineNormalizationThresholdSlider.Value;
            SaveSettingsToFile();
        }

        #endregion

        #region Advanced

        private void ToggleSwitchIsSpecialScreen_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Advanced.IsSpecialScreen = ToggleSwitchIsSpecialScreen.IsOn;
            TouchMultiplierSlider.Visibility = ToggleSwitchIsSpecialScreen.IsOn ? Visibility.Visible : Visibility.Collapsed;
            SaveSettingsToFile();
        }

        private void TouchMultiplierSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            Settings.Advanced.TouchMultiplier = e.NewValue;
            SaveSettingsToFile();
        }

        private void BorderCalculateMultiplier_TouchDown(object sender, TouchEventArgs e)
        {
            var args = e.GetTouchPoint(null).Bounds;
            double value;
            if (!Settings.Advanced.IsQuadIR) value = args.Width;
            else value = Math.Sqrt(args.Width * args.Height); //四边红外

            TextBlockShowCalculatedMultiplier.Text = (5 / (value * 1.1)).ToString();
        }

        private void ToggleSwitchEraserBindTouchMultiplier_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Advanced.EraserBindTouchMultiplier = ToggleSwitchEraserBindTouchMultiplier.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchIsQuadIR_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Advanced.IsQuadIR = ToggleSwitchIsQuadIR.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchIsLogEnabled_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Advanced.IsLogEnabled = ToggleSwitchIsLogEnabled.IsOn;
            SaveSettingsToFile();
        }

        private void ToggleSwitchDisableEdgeGesture_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Advanced.DisableEdgeGesture = ToggleSwitchDisableEdgeGesture.IsOn;
            EdgeGesturesUtils.DisableEdgeGestures(new WindowInteropHelper(this).Handle,
                Settings.Advanced.DisableEdgeGesture && Main_Grid.Background != Brushes.Transparent);
            SaveSettingsToFile();
        }

        #endregion

        public static void SaveSettingsToFile()
        {
            string text = JsonConvert.SerializeObject(Settings, Formatting.Indented);
            try
            {
                File.WriteAllText(App.RootPath + settingsFileName, text);
            }
            catch { }
        }

        private void SaveFloatBarPositionToFile()
        {
            try
            {
                File.WriteAllText(App.RootPath + positionFileName,
                    ViewboxFloatingBar.Margin.Left.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
                    ViewboxFloatingBar.Margin.Top.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            catch { }
        }

        private void SCManipulationBoundaryFeedback(object sender, ManipulationBoundaryFeedbackEventArgs e)
        {
            e.Handled = true;
        }

        private void HyperlinkWebsite_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://www.khyan.top/apps/Ink-Canvas-Plus");
        }

        private void HyperlinkQQGroup_Click(object sender, RoutedEventArgs e)
        {
            FlyoutQQGroup.Hide();
            Process.Start("https://qm.qq.com/q/I6OCRh38oU");
        }

        private void HyperlinkSource_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://github.com/clover-yan/Ink-Canvas-Plus");
        }

        #endregion

    }
}
