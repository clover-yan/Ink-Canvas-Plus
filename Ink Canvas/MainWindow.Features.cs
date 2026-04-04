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
    public partial class MainWindow : Window
    {
        #region Whiteboard Controls

        StrokeCollection[] strokeCollections = new StrokeCollection[101];
        bool[] whiteboadLastModeIsRedo = new bool[101];
        StrokeCollection lastTouchDownStrokeCollection = new StrokeCollection();

        int CurrentWhiteboardIndex = 1;
        int WhiteboardTotalCount = 1;
        TimeMachineHistory[][] TimeMachineHistories = new TimeMachineHistory[101][]; //最多99页，0用来存储非白板时的墨迹以便还原

        private void SaveStrokes(bool isBackupMain = false)
        {
            if (isBackupMain)
            {
                var timeMachineHistory = timeMachine.ExportTimeMachineHistory();
                TimeMachineHistories[0] = timeMachineHistory;
                timeMachine.ClearStrokeHistory();

            }
            else
            {
                var timeMachineHistory = timeMachine.ExportTimeMachineHistory();
                TimeMachineHistories[CurrentWhiteboardIndex] = timeMachineHistory;
                timeMachine.ClearStrokeHistory();
            }
        }

        private void ClearStrokes(bool isErasedByCode)
        {

            _currentCommitType = CommitReason.ClearingCanvas;
            if (isErasedByCode) _currentCommitType = CommitReason.CodeInput;
            inkCanvas.Strokes.Clear();
            _currentCommitType = CommitReason.UserInput;
        }

        private void RestoreStrokes(bool isBackupMain = false)
        {
            try
            {
                if (TimeMachineHistories[CurrentWhiteboardIndex] == null) return; //防止白板打开后不居中
                if (isBackupMain)
                {
                    timeMachine.ImportTimeMachineHistory(TimeMachineHistories[0]);
                    foreach (var item in TimeMachineHistories[0])
                    {
                        ApplyHistoryToCanvas(item);
                    }
                }
                else
                {
                    timeMachine.ImportTimeMachineHistory(TimeMachineHistories[CurrentWhiteboardIndex]);
                    foreach (var item in TimeMachineHistories[CurrentWhiteboardIndex])
                    {
                        ApplyHistoryToCanvas(item);
                    }
                }
            }
            catch { }
        }

        private void BtnWhiteBoardSwitchPrevious_Click(object sender, EventArgs e)
        {
            if (CurrentWhiteboardIndex <= 1) return;

            SaveStrokes();

            ClearStrokes(true);
            CurrentWhiteboardIndex--;

            RestoreStrokes();

            UpdateIndexInfoDisplay();
        }

        private void BtnWhiteBoardSwitchNext_Click(object sender, EventArgs e)
        {
            if (CurrentWhiteboardIndex >= WhiteboardTotalCount)
            {
                BtnWhiteBoardAdd_Click(sender, e);
                return;
            }
            if (Settings.Automation.IsAutoSaveStrokesAtClear && inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber)
            {
                SaveScreenShot(true);
                if (Settings.Automation.IsAutoSaveStrokesAtScreenshot) SaveInkCanvasStrokes(false);
            }
            SaveStrokes();


            ClearStrokes(true);
            CurrentWhiteboardIndex++;

            RestoreStrokes();

            UpdateIndexInfoDisplay();
        }

        private void BtnWhiteBoardAdd_Click(object sender, EventArgs e)
        {
            if (WhiteboardTotalCount >= 99) return;
            if (Settings.Automation.IsAutoSaveStrokesAtClear && inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber)
            {
                SaveScreenShot(true);
                if (Settings.Automation.IsAutoSaveStrokesAtScreenshot) SaveInkCanvasStrokes(false);
            }
            SaveStrokes();
            ClearStrokes(true);

            WhiteboardTotalCount++;
            CurrentWhiteboardIndex++;

            if (CurrentWhiteboardIndex != WhiteboardTotalCount)
            {
                for (int i = WhiteboardTotalCount; i > CurrentWhiteboardIndex; i--)
                {
                    TimeMachineHistories[i] = TimeMachineHistories[i - 1];
                }
            }

            UpdateIndexInfoDisplay();

            if (WhiteboardTotalCount >= 99) BtnWhiteBoardAdd.IsEnabled = false;
        }

        private void BtnWhiteBoardDelete_Click(object sender, RoutedEventArgs e)
        {
            ClearStrokes(true);

            if (CurrentWhiteboardIndex != WhiteboardTotalCount)
            {
                for (int i = CurrentWhiteboardIndex; i <= WhiteboardTotalCount; i++)
                {
                    TimeMachineHistories[i] = TimeMachineHistories[i + 1];
                }
            }
            else
            {
                CurrentWhiteboardIndex--;
            }

            WhiteboardTotalCount--;

            RestoreStrokes();

            UpdateIndexInfoDisplay();

            if (WhiteboardTotalCount < 99) BtnWhiteBoardAdd.IsEnabled = true;
        }

        private void UpdateIndexInfoDisplay()
        {
            TextBlockWhiteBoardIndexInfo.Text = string.Format("{0} / {1}", CurrentWhiteboardIndex, WhiteboardTotalCount);

            if (CurrentWhiteboardIndex == 1)
            {
                BtnWhiteBoardSwitchPrevious.IsEnabled = false;
            }
            else
            {
                BtnWhiteBoardSwitchPrevious.IsEnabled = true;
            }

            if (CurrentWhiteboardIndex == WhiteboardTotalCount)
            {
                BtnWhiteBoardSwitchNext.IsEnabled = false;
            }
            else
            {
                BtnWhiteBoardSwitchNext.IsEnabled = true;
            }

            if (WhiteboardTotalCount == 1)
            {
                BtnWhiteBoardDelete.IsEnabled = false;
            }
            else
            {
                BtnWhiteBoardDelete.IsEnabled = true;
            }
        }

        private void SetColors()
        {
            if (currentMode != 0 && !Settings.Canvas.UsingWhiteboard)
            {
                if (File.Exists(App.RootPath + "Colors\\Light.ini"))
                {
                    try
                    {
                        string[] lightColors = File.ReadAllLines(App.RootPath + "Colors\\Light.ini");
                        BtnColorRed.Background = new SolidColorBrush(StringToColor(lightColors[0]));
                        BtnColorGreen.Background = new SolidColorBrush(StringToColor(lightColors[1]));
                        BtnColorBlue.Background = new SolidColorBrush(StringToColor(lightColors[2]));
                        BtnColorYellow.Background = new SolidColorBrush(StringToColor(lightColors[3]));
                    }
                    catch (Exception) { ShowNotification("读取亮色画笔颜色配置文件时遇到问题"); }
                }
                else
                {
                    BtnColorRed.Background = new SolidColorBrush(StringToColor("#FFFF3333"));
                    BtnColorGreen.Background = new SolidColorBrush(StringToColor("#FF1ED760"));
                    BtnColorBlue.Background = new SolidColorBrush(StringToColor("#FF239AD6"));
                    BtnColorYellow.Background = new SolidColorBrush(StringToColor("#FFFFC000"));
                }
            }
            else
            {
                if (File.Exists(App.RootPath + "Colors\\Dark.ini"))
                {
                    try
                    {
                        string[] darkColors = File.ReadAllLines(App.RootPath + "Colors\\Dark.ini");
                        BtnColorRed.Background = new SolidColorBrush(StringToColor(darkColors[0]));
                        BtnColorGreen.Background = new SolidColorBrush(StringToColor(darkColors[1]));
                        BtnColorBlue.Background = new SolidColorBrush(StringToColor(darkColors[2]));
                        BtnColorYellow.Background = new SolidColorBrush(StringToColor(darkColors[3]));
                    }
                    catch (Exception) { ShowNotification("读取深色画笔颜色配置文件时遇到问题"); }
                }
                else
                {
                    BtnColorRed.Background = new SolidColorBrush(Colors.Red);
                    BtnColorGreen.Background = new SolidColorBrush(StringToColor("#FF169141"));
                    BtnColorBlue.Background = new SolidColorBrush(StringToColor("#FF239AD6"));
                    BtnColorYellow.Background = new SolidColorBrush(StringToColor("#FFF38B00"));
                }
            }
        }

        #endregion Whiteboard Controls

        #region Simulate Pen Pressure & Ink To Shape

        StrokeCollection newStrokes = new StrokeCollection();
        List<Circle> circles = new List<Circle>();

        //此函数中的所有代码版权所有 WXRIW，在其他项目中使用前必须提前联系（wxriw@outlook.com），谢谢！
        private void inkCanvas_StrokeCollected(object sender, InkCanvasStrokeCollectedEventArgs e)
        {
            try
            {
                inkCanvas.Opacity = 1;
                if (Settings.InkToShape.IsInkToShapeEnabled && !Environment.Is64BitProcess)
                {
                    void InkToShapeProcess()
                    {
                        try
                        {
                            newStrokes.Add(e.Stroke);
                            if (newStrokes.Count > 4) newStrokes.RemoveAt(0);
                            for (int i = 0; i < newStrokes.Count; i++)
                            {
                                if (!inkCanvas.Strokes.Contains(newStrokes[i])) newStrokes.RemoveAt(i--);
                            }
                            for (int i = 0; i < circles.Count; i++)
                            {
                                if (!inkCanvas.Strokes.Contains(circles[i].Stroke)) circles.RemoveAt(i);
                            }
                            var strokeReco = new StrokeCollection();
                            var result = InkRecognizeHelper.RecognizeShape(newStrokes);
                            for (int i = newStrokes.Count - 1; i >= 0; i--)
                            {
                                strokeReco.Add(newStrokes[i]);
                                var newResult = InkRecognizeHelper.RecognizeShape(strokeReco);
                                if (newResult.InkDrawingNode.GetShapeName() == "Circle" || newResult.InkDrawingNode.GetShapeName() == "Ellipse")
                                {
                                    result = newResult;
                                    break;
                                }
                                //Label.Visibility = Visibility.Visible;
                                Label.Content = circles.Count.ToString() + "\n" + newResult.InkDrawingNode.GetShapeName();
                            }
                            if (result.InkDrawingNode.GetShapeName() == "Circle")
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                if (shape.Width > 75)
                                {
                                    foreach (Circle circle in circles)
                                    {
                                        //判断是否画同心圆
                                        if (Math.Abs(result.Centroid.X - circle.Centroid.X) / shape.Width < 0.12 &&
                                            Math.Abs(result.Centroid.Y - circle.Centroid.Y) / shape.Width < 0.12)
                                        {
                                            result.Centroid = circle.Centroid;
                                            break;
                                        }
                                        else
                                        {
                                            double d = (result.Centroid.X - circle.Centroid.X) * (result.Centroid.X - circle.Centroid.X) +
                                               (result.Centroid.Y - circle.Centroid.Y) * (result.Centroid.Y - circle.Centroid.Y);
                                            d = Math.Sqrt(d);
                                            //判断是否画外切圆
                                            double x = shape.Width / 2.0 + circle.R - d;
                                            if (Math.Abs(x) / shape.Width < 0.1)
                                            {
                                                double sinTheta = (result.Centroid.Y - circle.Centroid.Y) / d;
                                                double cosTheta = (result.Centroid.X - circle.Centroid.X) / d;
                                                double newX = result.Centroid.X + x * cosTheta;
                                                double newY = result.Centroid.Y + x * sinTheta;
                                                result.Centroid = new Point(newX, newY);
                                            }
                                            //判断是否画外切圆
                                            x = Math.Abs(circle.R - shape.Width / 2.0) - d;
                                            if (Math.Abs(x) / shape.Width < 0.1)
                                            {
                                                double sinTheta = (result.Centroid.Y - circle.Centroid.Y) / d;
                                                double cosTheta = (result.Centroid.X - circle.Centroid.X) / d;
                                                double newX = result.Centroid.X + x * cosTheta;
                                                double newY = result.Centroid.Y + x * sinTheta;
                                                result.Centroid = new Point(newX, newY);
                                            }
                                        }
                                    }

                                    Point iniP = new Point(result.Centroid.X - shape.Width / 2, result.Centroid.Y - shape.Height / 2);
                                    Point endP = new Point(result.Centroid.X + shape.Width / 2, result.Centroid.Y + shape.Height / 2);
                                    var pointList = GenerateEllipseGeometry(iniP, endP);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(point)
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };
                                    circles.Add(new Circle(result.Centroid, shape.Width / 2.0, stroke));
                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else if (result.InkDrawingNode.GetShapeName().Contains("Ellipse"))
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                //var shape1 = result.InkDrawingNode.GetShape();
                                //shape1.Fill = Brushes.Gray;
                                //Canvas.Children.Add(shape1);
                                var p = result.InkDrawingNode.HotPoints;
                                double a = GetDistance(p[0], p[2]) / 2; //长半轴
                                double b = GetDistance(p[1], p[3]) / 2; //短半轴
                                if (a < b)
                                {
                                    double t = a;
                                    a = b;
                                    b = t;
                                }

                                result.Centroid = new Point((p[0].X + p[2].X) / 2, (p[0].Y + p[2].Y) / 2);
                                bool needRotation = true;

                                if (shape.Width > 75 || shape.Height > 75 && p.Count == 4)
                                {
                                    Point iniP = new Point(result.Centroid.X - shape.Width / 2, result.Centroid.Y - shape.Height / 2);
                                    Point endP = new Point(result.Centroid.X + shape.Width / 2, result.Centroid.Y + shape.Height / 2);

                                    foreach (Circle circle in circles)
                                    {
                                        //判断是否画同心椭圆
                                        if (Math.Abs(result.Centroid.X - circle.Centroid.X) / a < 0.2 &&
                                            Math.Abs(result.Centroid.Y - circle.Centroid.Y) / a < 0.2)
                                        {
                                            result.Centroid = circle.Centroid;
                                            iniP = new Point(result.Centroid.X - shape.Width / 2, result.Centroid.Y - shape.Height / 2);
                                            endP = new Point(result.Centroid.X + shape.Width / 2, result.Centroid.Y + shape.Height / 2);

                                            //再判断是否与圆相切
                                            if (Math.Abs(a - circle.R) / a < 0.2)
                                            {
                                                if (shape.Width >= shape.Height)
                                                {
                                                    iniP.X = result.Centroid.X - circle.R;
                                                    endP.X = result.Centroid.X + circle.R;
                                                    iniP.Y = result.Centroid.Y - b;
                                                    endP.Y = result.Centroid.Y + b;
                                                }
                                                else
                                                {
                                                    iniP.Y = result.Centroid.Y - circle.R;
                                                    endP.Y = result.Centroid.Y + circle.R;
                                                    iniP.X = result.Centroid.X - a;
                                                    endP.X = result.Centroid.X + a;
                                                }
                                            }
                                            break;
                                        }
                                        else if (Math.Abs(result.Centroid.X - circle.Centroid.X) / a < 0.2)
                                        {
                                            double sinTheta = Math.Abs(circle.Centroid.Y - result.Centroid.Y) / circle.R;
                                            double cosTheta = Math.Sqrt(1 - sinTheta * sinTheta);
                                            double newA = circle.R * cosTheta;
                                            if (circle.R * sinTheta / circle.R < 0.9 && a / b > 2 && Math.Abs(newA - a) / newA < 0.3)
                                            {
                                                iniP.X = circle.Centroid.X - newA;
                                                endP.X = circle.Centroid.X + newA;
                                                iniP.Y = result.Centroid.Y - newA / 5;
                                                endP.Y = result.Centroid.Y + newA / 5;

                                                double topB = endP.Y - iniP.Y;

                                                SetNewBackupOfStroke();
                                                _currentCommitType = CommitReason.ShapeRecognition;
                                                inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                                newStrokes = new StrokeCollection();

                                                var _pointList = GenerateEllipseGeometry(iniP, endP, false, true);
                                                var _point = new StylusPointCollection(_pointList);
                                                var _stroke = new Stroke(_point)
                                                {
                                                    DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                                };
                                                var _dashedLineStroke = GenerateDashedLineEllipseStrokeCollection(iniP, endP, true, false);
                                                StrokeCollection strokes = new StrokeCollection()
                                                {
                                                    _stroke,
                                                    _dashedLineStroke
                                                };
                                                inkCanvas.Strokes.Add(strokes);
                                                _currentCommitType = CommitReason.UserInput;
                                                return;
                                            }
                                        }
                                        else if (Math.Abs(result.Centroid.Y - circle.Centroid.Y) / a < 0.2)
                                        {
                                            double cosTheta = Math.Abs(circle.Centroid.X - result.Centroid.X) / circle.R;
                                            double sinTheta = Math.Sqrt(1 - cosTheta * cosTheta);
                                            double newA = circle.R * sinTheta;
                                            if (circle.R * sinTheta / circle.R < 0.9 && a / b > 2 && Math.Abs(newA - a) / newA < 0.3)
                                            {
                                                iniP.X = result.Centroid.X - newA / 5;
                                                endP.X = result.Centroid.X + newA / 5;
                                                iniP.Y = circle.Centroid.Y - newA;
                                                endP.Y = circle.Centroid.Y + newA;
                                                needRotation = false;
                                            }
                                        }
                                    }

                                    //纠正垂直与水平关系
                                    var newPoints = FixPointsDirection(p[0], p[2]);
                                    p[0] = newPoints[0];
                                    p[2] = newPoints[1];
                                    newPoints = FixPointsDirection(p[1], p[3]);
                                    p[1] = newPoints[0];
                                    p[3] = newPoints[1];

                                    var pointList = GenerateEllipseGeometry(iniP, endP);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(point)
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };

                                    if (needRotation)
                                    {
                                        Matrix m = new Matrix();
                                        FrameworkElement fe = e.Source as FrameworkElement;
                                        double tanTheta = (p[2].Y - p[0].Y) / (p[2].X - p[0].X);
                                        double theta = Math.Atan(tanTheta);
                                        m.RotateAt(theta * 180.0 / Math.PI, result.Centroid.X, result.Centroid.Y);
                                        stroke.Transform(m, false);
                                    }

                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else if (result.InkDrawingNode.GetShapeName().Contains("Triangle"))
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                var p = result.InkDrawingNode.HotPoints;
                                if ((Math.Max(Math.Max(p[0].X, p[1].X), p[2].X) - Math.Min(Math.Min(p[0].X, p[1].X), p[2].X) >= 100 ||
                                    Math.Max(Math.Max(p[0].Y, p[1].Y), p[2].Y) - Math.Min(Math.Min(p[0].Y, p[1].Y), p[2].Y) >= 100) && result.InkDrawingNode.HotPoints.Count == 3)
                                {
                                    //纠正垂直与水平关系
                                    var newPoints = FixPointsDirection(p[0], p[1]);
                                    p[0] = newPoints[0];
                                    p[1] = newPoints[1];
                                    newPoints = FixPointsDirection(p[0], p[2]);
                                    p[0] = newPoints[0];
                                    p[2] = newPoints[1];
                                    newPoints = FixPointsDirection(p[1], p[2]);
                                    p[1] = newPoints[0];
                                    p[2] = newPoints[1];

                                    var pointList = p.ToList();
                                    //pointList.Add(p[0]);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(GenerateFakePressureTriangle(point))
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };
                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else if (result.InkDrawingNode.GetShapeName().Contains("Rectangle") ||
                                     result.InkDrawingNode.GetShapeName().Contains("Diamond") ||
                                     result.InkDrawingNode.GetShapeName().Contains("Parallelogram") ||
                                     result.InkDrawingNode.GetShapeName().Contains("Square"))
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                var p = result.InkDrawingNode.HotPoints;
                                if ((Math.Max(Math.Max(Math.Max(p[0].X, p[1].X), p[2].X), p[3].X) - Math.Min(Math.Min(Math.Min(p[0].X, p[1].X), p[2].X), p[3].X) >= 100 ||
                                    Math.Max(Math.Max(Math.Max(p[0].Y, p[1].Y), p[2].Y), p[3].Y) - Math.Min(Math.Min(Math.Min(p[0].Y, p[1].Y), p[2].Y), p[3].Y) >= 100) && result.InkDrawingNode.HotPoints.Count == 4)
                                {
                                    //纠正垂直与水平关系
                                    var newPoints = FixPointsDirection(p[0], p[1]);
                                    p[0] = newPoints[0];
                                    p[1] = newPoints[1];
                                    newPoints = FixPointsDirection(p[1], p[2]);
                                    p[1] = newPoints[0];
                                    p[2] = newPoints[1];
                                    newPoints = FixPointsDirection(p[2], p[3]);
                                    p[2] = newPoints[0];
                                    p[3] = newPoints[1];
                                    newPoints = FixPointsDirection(p[3], p[0]);
                                    p[3] = newPoints[0];
                                    p[0] = newPoints[1];

                                    var pointList = p.ToList();
                                    pointList.Add(p[0]);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(GenerateFakePressureRectangle(point))
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };
                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else
                            {
                                // 检查是否是直线
                                Stroke strokeToBeChecked = e.Stroke;
                                var points = strokeToBeChecked.StylusPoints.Select(p => p.ToPoint()).ToList();
                                if (points.Count >= 10)
                                {
                                    // 使用总最小二乘法(TLS/PCA)进行直线拟合
                                    int n = points.Count - 8;
                                    List<Point> filteredPoints = new List<Point>();

                                    // 收集过滤后的点（跳过前 4 个和后 4 个点，用于计算直线方向）
                                    for (int i = 4; i < n + 4; i++)
                                    {
                                        filteredPoints.Add(points[i]);
                                    }

                                    // 计算中心点（使用过滤后的点）
                                    double centerX = 0, centerY = 0;
                                    foreach (Point p in filteredPoints)
                                    {
                                        centerX += p.X;
                                        centerY += p.Y;
                                    }
                                    centerX /= filteredPoints.Count;
                                    centerY /= filteredPoints.Count;

                                    // 计算协方差矩阵（使用过滤后的点）
                                    double covXX = 0, covYY = 0, covXY = 0;
                                    foreach (Point p in filteredPoints)
                                    {
                                        double dx = p.X - centerX;
                                        double dy = p.Y - centerY;
                                        covXX += dx * dx;
                                        covYY += dy * dy;
                                        covXY += dx * dy;
                                    }

                                    // 计算特征值和特征向量
                                    double trace = covXX + covYY;
                                    double determinant = covXX * covYY - covXY * covXY;
                                    double discriminant = Math.Sqrt(trace * trace - 4 * determinant);

                                    double eigenvalue1 = (trace + discriminant) / 2;
                                    double eigenvalue2 = (trace - discriminant) / 2;

                                    // 最大特征值对应的特征向量即为直线方向
                                    double directionX, directionY;
                                    if (Math.Abs(covXY) > 1e-10)
                                    {
                                        directionX = covXY;
                                        directionY = eigenvalue1 - covXX;
                                        // 归一化
                                        double length = Math.Sqrt(directionX * directionX + directionY * directionY);
                                        directionX /= length;
                                        directionY /= length;
                                    }
                                    else
                                    {
                                        // 如果协方差为 0，则是水平或垂直直线
                                        directionX = (covXX >= covYY) ? 1 : 0;
                                        directionY = (covXX >= covYY) ? 0 : 1;
                                    }

                                    // 计算解释方差比例（拟合优度）
                                    double totalVariance = eigenvalue1 + eigenvalue2;
                                    double explainedVarianceRatio = (totalVariance > 1e-10) ?
                                        Math.Max(eigenvalue1, eigenvalue2) / totalVariance : 1d;

                                    // 使用所有点计算端点
                                    double minProjection = double.MaxValue;
                                    double maxProjection = double.MinValue;

                                    // 计算所有点在直线方向上的投影
                                    foreach (Point p in points)
                                    {
                                        // 相对于过滤点中心的投影
                                        double projection = (p.X - centerX) * directionX + (p.Y - centerY) * directionY;
                                        minProjection = Math.Min(minProjection, projection);
                                        maxProjection = Math.Max(maxProjection, projection);
                                    }

                                    // 计算端点坐标
                                    Point endpoint1 = new Point(
                                        centerX + minProjection * directionX,
                                        centerY + minProjection * directionY
                                    );

                                    Point endpoint2 = new Point(
                                        centerX + maxProjection * directionX,
                                        centerY + maxProjection * directionY
                                    );

                                    // 使用解释方差比例作为判断条件
                                    if (explainedVarianceRatio > 0.998 + Settings.InkToShape.LineNormalizationThreshold / 500)
                                    {
                                        var pointList = new List<Point> { endpoint1, endpoint2 };
                                        var point = new StylusPointCollection(pointList);
                                        var stroke = new Stroke(point)
                                        {
                                            DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                        };
                                        SetNewBackupOfStroke();
                                        _currentCommitType = CommitReason.ShapeRecognition;
                                        inkCanvas.Strokes.Remove(strokeToBeChecked);
                                        inkCanvas.Strokes.Add(stroke);
                                        _currentCommitType = CommitReason.UserInput;
                                        GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                        newStrokes = new StrokeCollection();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogHelper.WriteLogToFile($"{ex.Message}");
                        }
                    }
                    InkToShapeProcess();
                }

                // 检查是否是压感笔书写
                foreach (StylusPoint stylusPoint in e.Stroke.StylusPoints)
                {
                    if (stylusPoint.PressureFactor != 0.5 && stylusPoint.PressureFactor != 0)
                    {
                        return;
                    }
                }


                try
                {
                    if (e.Stroke.StylusPoints.Count > 3)
                    {
                        Random random = new Random();
                        double _speed = GetPointSpeed(e.Stroke.StylusPoints[random.Next(0, e.Stroke.StylusPoints.Count - 1)].ToPoint(), e.Stroke.StylusPoints[random.Next(0, e.Stroke.StylusPoints.Count - 1)].ToPoint(), e.Stroke.StylusPoints[random.Next(0, e.Stroke.StylusPoints.Count - 1)].ToPoint());

                        RandWindow.randSeed = (int)(_speed * 100000 * 1000);
                    }
                }
                catch { }

                switch (Settings.Canvas.InkStyle)
                {
                    case 1:
                        try
                        {
                            StylusPointCollection stylusPoints = new StylusPointCollection();
                            int n = e.Stroke.StylusPoints.Count - 1;
                            string s = "";

                            for (int i = 0; i <= n; i++)
                            {
                                double speed = GetPointSpeed(e.Stroke.StylusPoints[Math.Max(i - 1, 0)].ToPoint(), e.Stroke.StylusPoints[i].ToPoint(), e.Stroke.StylusPoints[Math.Min(i + 1, n)].ToPoint());
                                s += speed.ToString() + "\t";
                                StylusPoint point = new StylusPoint();
                                if (speed >= 0.25)
                                {
                                    point.PressureFactor = (float)(0.5 - 0.3 * (Math.Min(speed, 1.5) - 0.3) / 1.2);
                                }
                                else if (speed >= 0.05)
                                {
                                    point.PressureFactor = (float)0.5;
                                }
                                else
                                {
                                    point.PressureFactor = (float)(0.5 + 0.4 * (0.05 - speed) / 0.05);
                                }
                                point.X = e.Stroke.StylusPoints[i].X;
                                point.Y = e.Stroke.StylusPoints[i].Y;
                                stylusPoints.Add(point);
                            }
                            //Label.Visibility = Visibility.Visible;
                            //Label.Content = s;
                            e.Stroke.StylusPoints = stylusPoints;
                        }
                        catch
                        {

                        }
                        break;
                    case 0:
                        try
                        {
                            StylusPointCollection stylusPoints = new StylusPointCollection();
                            int n = e.Stroke.StylusPoints.Count - 1;
                            double pressure = 0.1;
                            int x = 10;
                            if (n == 1) return;
                            if (n >= x)
                            {
                                for (int i = 0; i < n - x; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)0.5;
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                                for (int i = n - x; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)((0.5 - pressure) * (n - i) / x + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            else
                            {
                                for (int i = 0; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)(0.4 * (n - i) / n + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            e.Stroke.StylusPoints = stylusPoints;
                        }
                        catch
                        {

                        }
                        break;
                    case 3: //根据 mode == 0 改写，目前暂未完成
                        try
                        {
                            StylusPointCollection stylusPoints = new StylusPointCollection();
                            int n = e.Stroke.StylusPoints.Count - 1;
                            double pressure = 0.1;
                            int x = 8;
                            if (lastTouchDownTime < lastTouchUpTime)
                            {
                                double k = (lastTouchUpTime - lastTouchDownTime) / (n + 1); // 每个点之间间隔 k 毫秒
                                x = (int)(1000 / k); // 取 1000 ms 内的点
                            }

                            if (n >= x)
                            {
                                for (int i = 0; i < n - x; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)0.5;
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                                for (int i = n - x; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)((0.5 - pressure) * (n - i) / x + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            else
                            {
                                for (int i = 0; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)(0.4 * (n - i) / n + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            e.Stroke.StylusPoints = stylusPoints;
                        }
                        catch
                        {

                        }
                        break;
                }
            }
            catch { }
        }

        private void SetNewBackupOfStroke()
        {
            lastTouchDownStrokeCollection = inkCanvas.Strokes.Clone();
            int whiteboardIndex = CurrentWhiteboardIndex;
            if (currentMode == 0)
            {
                whiteboardIndex = 0;
            }
            strokeCollections[whiteboardIndex] = lastTouchDownStrokeCollection;
        }

        public double GetDistance(Point point1, Point point2)
        {
            return Math.Sqrt((point1.X - point2.X) * (point1.X - point2.X) + (point1.Y - point2.Y) * (point1.Y - point2.Y));
        }

        public double GetPointSpeed(Point point1, Point point2, Point point3)
        {
            return (Math.Sqrt((point1.X - point2.X) * (point1.X - point2.X) + (point1.Y - point2.Y) * (point1.Y - point2.Y))
                + Math.Sqrt((point3.X - point2.X) * (point3.X - point2.X) + (point3.Y - point2.Y) * (point3.Y - point2.Y)))
                / 20;
        }

        public Point[] FixPointsDirection(Point p1, Point p2)
        {
            if (Math.Abs(p1.X - p2.X) / Math.Abs(p1.Y - p2.Y) > 8)
            {
                //水平
                double x = Math.Abs(p1.Y - p2.Y) / 2;
                if (p1.Y > p2.Y)
                {
                    p1.Y -= x;
                    p2.Y += x;
                }
                else
                {
                    p1.Y += x;
                    p2.Y -= x;
                }
            }
            else if (Math.Abs(p1.Y - p2.Y) / Math.Abs(p1.X - p2.X) > 8)
            {
                //垂直
                double x = Math.Abs(p1.X - p2.X) / 2;
                if (p1.X > p2.X)
                {
                    p1.X -= x;
                    p2.X += x;
                }
                else
                {
                    p1.X += x;
                    p2.X -= x;
                }
            }

            return new Point[2] { p1, p2 };
        }

        public StylusPointCollection GenerateFakePressureTriangle(StylusPointCollection points)
        {
            var newPoint = new StylusPointCollection();
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            var cPoint = GetCenterPoint(points[0], points[1]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            cPoint = GetCenterPoint(points[1], points[2]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            cPoint = GetCenterPoint(points[2], points[0]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            return newPoint;
        }

        public StylusPointCollection GenerateFakePressureRectangle(StylusPointCollection points)
        {
            var newPoint = new StylusPointCollection();
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            var cPoint = GetCenterPoint(points[0], points[1]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            cPoint = GetCenterPoint(points[1], points[2]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            cPoint = GetCenterPoint(points[2], points[3]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[3].X, points[3].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[3].X, points[3].Y, (float)0.4));
            cPoint = GetCenterPoint(points[3], points[0]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            return newPoint;
        }

        public Point GetCenterPoint(Point point1, Point point2)
        {
            return new Point((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2);
        }

        public StylusPoint GetCenterPoint(StylusPoint point1, StylusPoint point2)
        {
            return new StylusPoint((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2);
        }

        #endregion

        #region Functions

        /// <summary>
        /// 传入域名返回对应的IP 
        /// </summary>
        /// <param name="domainName">域名</param>
        /// <returns></returns>
        public static string GetIp(string domainName)
        {
            domainName = domainName.Replace("http://", "").Replace("https://", "");
            IPHostEntry hostEntry = Dns.GetHostEntry(domainName);
            IPEndPoint ipEndPoint = new IPEndPoint(hostEntry.AddressList[0], 0);
            return ipEndPoint.Address.ToString();
        }

        public static string GetWebClient(string url)
        {
            HttpWebRequest myrq = (HttpWebRequest)WebRequest.Create(url);

            myrq.Proxy = null;
            myrq.KeepAlive = false;
            myrq.Timeout = 30 * 1000;
            myrq.Method = "Get";
            myrq.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            myrq.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 UBrowser/6.2.4098.3 Safari/537.36";

            HttpWebResponse myrp;
            try
            {
                myrp = (HttpWebResponse)myrq.GetResponse();
            }
            catch (WebException ex)
            {
                myrp = (HttpWebResponse)ex.Response;
            }

            if (myrp?.StatusCode != HttpStatusCode.OK)
            {
                return "null";
            }

            using (StreamReader sr = new StreamReader(myrp.GetResponseStream()))
            {
                return sr.ReadToEnd();
            }
        }

        #region 开机自启
        /// <summary>
        /// 开机自启创建
        /// </summary>
        /// <param name="exeName">程序名称</param>
        /// <returns></returns>
        public static bool StartAutomaticallyCreate(string exeName)
        {
            try
            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + exeName + ".lnk");
                //设置快捷方式的目标所在的位置(源程序完整路径)
                shortcut.TargetPath = System.Windows.Forms.Application.ExecutablePath;
                //应用程序的工作目录
                //当用户没有指定一个具体的目录时，快捷方式的目标应用程序将使用该属性所指定的目录来装载或保存文件。
                shortcut.WorkingDirectory = System.Environment.CurrentDirectory;
                //目标应用程序窗口类型(1.Normal window普通窗口,3.Maximized最大化窗口,7.Minimized最小化)
                shortcut.WindowStyle = 1;
                //快捷方式的描述
                shortcut.Description = exeName + "_Ink";
                //设置快捷键(如果有必要的话.)
                //shortcut.Hotkey = "CTRL+ALT+D";
                shortcut.Save();
                return true;
            }
            catch (Exception) { }
            return false;
        }

        /// <summary>
        /// 开机自启删除
        /// </summary>
        /// <param name="exeName">程序名称</param>
        /// <returns></returns>
        public static bool StartAutomaticallyDel(string exeName)
        {
            try
            {
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + exeName + ".lnk");
                return true;
            }
            catch (Exception) { }
            return false;
        }
        #endregion

        #region Auto Theme

        Color FloatBarForegroundColor = Color.FromRgb(102, 102, 102);
        private void SetTheme(string theme, int favor)
        {
            if (favor == 1)
            {
                ResourceDictionary rd1 = new ResourceDictionary() { Source = new Uri($"Resources/Styles/{theme}Light.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd1);

                ResourceDictionary rd2 = new ResourceDictionary() { Source = new Uri("Resources/DrawShapeImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd2);

                ResourceDictionary rd3 = new ResourceDictionary() { Source = new Uri("Resources/SeewoImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd3);

                ResourceDictionary rd4 = new ResourceDictionary() { Source = new Uri("Resources/IconImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd4);

                ThemeManager.SetRequestedTheme(window, ElementTheme.Light);

                FloatBarForegroundColor = (Color)Application.Current.FindResource("FloatBarForegroundColor");
            }
            else if (favor == 2)
            {
                ResourceDictionary rd1 = new ResourceDictionary() { Source = new Uri($"Resources/Styles/{theme}Dark.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd1);

                ResourceDictionary rd2 = new ResourceDictionary() { Source = new Uri("Resources/DrawShapeImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd2);

                ResourceDictionary rd3 = new ResourceDictionary() { Source = new Uri("Resources/SeewoImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd3);

                ResourceDictionary rd4 = new ResourceDictionary() { Source = new Uri("Resources/IconImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd4);

                ThemeManager.SetRequestedTheme(window, ElementTheme.Dark);

                FloatBarForegroundColor = (Color)Application.Current.FindResource("FloatBarForegroundColor");
            }

            SymbolIconSelect.Foreground = new SolidColorBrush(FloatBarForegroundColor);
            SymbolIconDelete.Foreground = new SolidColorBrush(FloatBarForegroundColor);
        }

        private void SystemEvents_UserPreferenceChanged(object sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
        {
            int favor = 0;
            switch (Settings.Appearance.Theme % 3)
            {
                case 0:
                    favor = 1;
                    break;
                case 1:
                    favor = 2;
                    break;
                case 2:
                    if (IsSystemThemeLight()) favor = 1;
                    else favor = 2;
                    break;
            }
            switch (Settings.Appearance.Theme / 3)
            {
                case 0:
                    SetTheme("", favor);
                    break;
                case 1:
                    SetTheme("Green", favor);
                    break;
                case 2:
                    SetTheme("Pink", favor);
                    break;
                case 3:
                    SetTheme("Red", favor);
                    break;
            }
        }

        private bool IsSystemThemeLight()
        {
            bool light = false;
            try
            {
                RegistryKey registryKey = Registry.CurrentUser;
                RegistryKey themeKey = registryKey.OpenSubKey("software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize");
                int keyValue = 0;
                if (themeKey != null)
                {
                    keyValue = (int)themeKey.GetValue("SystemUsesLightTheme");
                }
                if (keyValue == 1) light = true;
            }
            catch { }
            return light;
        }
        #endregion

        #endregion Functions

        #region Screenshot

        private void BtnScreenshot_Click(object sender, RoutedEventArgs e)
        {
            bool isHideNotification = false;
            if (sender is bool) isHideNotification = (bool)sender;

            GridNotifications.Visibility = Visibility.Collapsed;

            new Thread(new ThreadStart(() =>
            {
                Thread.Sleep(20);
                try
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                            SaveScreenShot(isHideNotification, $"{pptName}/{previousSlideID}_{DateTime.Now:HH-mm-ss}");
                        else
                            SaveScreenShot(isHideNotification);
                    });
                }
                catch
                {
                    if (!isHideNotification)
                    {
                        ShowNotification("截图保存失败");
                    }
                }

                try
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (inkCanvas.Visibility != Visibility.Visible || inkCanvas.Strokes.Count == 0 || !Settings.Automation.IsAutoSaveStrokesAtScreenshot) return;
                        SaveInkCanvasStrokes(false);
                    });
                }
                catch { }

                if (isHideNotification)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        BtnClear_Click(BtnClear, null);
                    });
                }
            })).Start();
        }

        private void SaveScreenShot(bool isHideNotification, string fileName = null)
        {
            var size = System.Windows.Forms.SystemInformation.PrimaryMonitorSize;
            var rc = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new System.Drawing.Size(size.Width, size.Height));
            var bitmap = new System.Drawing.Bitmap(rc.Width, rc.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (System.Drawing.Graphics memoryGrahics = System.Drawing.Graphics.FromImage(bitmap))
            {
                memoryGrahics.CopyFromScreen(rc.X, rc.Y, 0, 0, rc.Size, System.Drawing.CopyPixelOperation.SourceCopy);
            }

            if (Settings.Automation.IsSaveScreenshotsInDateFolders)
            {
                if (string.IsNullOrWhiteSpace(fileName))
                    fileName = DateTime.Now.ToString("HH-mm-ss");
                var savePath =
                    $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)}\Ink Canvas Screenshots\{DateTime.Now.Date:yyyyMMdd}\{fileName}.png";


                if (!Directory.Exists(Path.GetDirectoryName(savePath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(savePath));
                }

                bitmap.Save(savePath, ImageFormat.Png);

                if (!isHideNotification)
                {
                    ShowNotification("截图成功保存至 " + savePath);
                }
            }
            else
            {
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + @"\Ink Canvas Screenshots"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) +
                                              @"\Ink Canvas Screenshots");
                }

                bitmap.Save(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) +
                            @"\Ink Canvas Screenshots\" + DateTime.Now.ToString("u").Replace(':', '-') + ".png", ImageFormat.Png);

                if (!isHideNotification)
                {
                    ShowNotification("截图成功保存至 " + Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) +
                                     @"\Ink Canvas Screenshots\" + DateTime.Now.ToString("u").Replace(':', '-') + ".png");
                }
            }
        }

        #endregion

        #region Notification

        int lastNotificationShowTime = 0;
        int notificationShowTime = 2500;

        public static void ShowNewMessage(string notice, bool isShowImmediately = true)
        {
            (Application.Current?.Windows.Cast<Window>().FirstOrDefault(window => window is MainWindow) as MainWindow)?.ShowNotification(notice, isShowImmediately);
        }

        public void ShowNotification(string notice, bool isShowImmediately = true)
        {
            lastNotificationShowTime = Environment.TickCount;

            GridNotifications.Visibility = Visibility.Visible;
            //GridNotifications.Opacity = 1;
            TextBlockNotice.Text = notice;

            new Thread(new ThreadStart(() =>
            {
                Thread.Sleep(notificationShowTime + 200);
                if (Environment.TickCount - lastNotificationShowTime >= notificationShowTime)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        GridNotifications.Visibility = Visibility.Collapsed;
                        //DoubleAnimation daV = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromSeconds(0.15)));
                        //GridNotifications.BeginAnimation(UIElement.OpacityProperty, daV);

                        //new Thread(new ThreadStart(() => {
                        //    Thread.Sleep(200);
                        //    Application.Current.Dispatcher.Invoke(() =>
                        //    {
                        //        if (GridNotifications.Opacity == 0)
                        //        {
                        //            GridNotifications.Visibility = Visibility.Collapsed;
                        //            GridNotifications.Opacity = 1;
                        //        }
                        //    });
                        //})).Start();
                    });
                }
            })).Start();
        }

        private void AppendNotification(string notice)
        {
            TextBlockNotice.Text = TextBlockNotice.Text + Environment.NewLine + notice;
        }

        #endregion

        #region Tools

        private void BtnTools_Click(object sender, RoutedEventArgs e)
        {
            if (StackPanelToolButtons.Visibility == Visibility.Visible)
            {
                StackPanelToolButtons.Visibility = Visibility.Collapsed;
            }
            else
            {
                StackPanelToolButtons.Visibility = Visibility.Visible;
            }
        }

        private void BtnCountdownTimer_Click(object sender, RoutedEventArgs e)
        {
            StackPanelToolButtons.Visibility = Visibility.Collapsed;
            new CountdownTimerWindow().Show();
        }

        private void BtnRand_Click(object sender, RoutedEventArgs e)
        {
            StackPanelToolButtons.Visibility = Visibility.Collapsed;
            new RandWindow().Show();
        }

        #endregion Tools

        #region Float Bar

        private void HideSubPanels()
        {
            BorderClearInDelete.Visibility = Visibility.Collapsed;
            BorderTools.Visibility = Visibility.Collapsed;
        }


        private void BorderPenColorBlack_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorBlack_Click(BtnColorBlack, null);
            HideSubPanels();
        }

        private void BorderPenColorRed_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorRed_Click(BtnColorRed, null);
            HideSubPanels();
        }

        private void BorderPenColorGreen_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorGreen_Click(BtnColorGreen, null);
            HideSubPanels();
        }

        private void BorderPenColorBlue_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorBlue_Click(BtnColorBlue, null);
            HideSubPanels();
        }

        private void BorderPenColorYellow_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorYellow_Click(BtnColorYellow, null);
            HideSubPanels();
        }

        private void BorderPenColorWhite_MouseUp(object sender, MouseButtonEventArgs e)
        {
            inkCanvas.DefaultDrawingAttributes.Color = StringToColor("#FFFEFEFE");
            inkColor = 5;
            ColorSwitchCheck();
            HideSubPanels();
        }

        private void SymbolIconUndo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnUndo_Click(BtnUndo, null);
            HideSubPanels();
        }

        private void SymbolIconRedo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnRedo_Click(BtnRedo, null);
            HideSubPanels();
        }

        private void SymbolIconCursor_Click(object sender, RoutedEventArgs e)
        {
            if (currentMode != 0)
            {
                ImageBlackboard_MouseUp(null, null);
            }
            else
            {
                var width = ViewboxFloatingBar.ActualWidth;
                using (Dispatcher.DisableProcessing())
                {
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);

                    ViewboxFloatingBar.UpdateLayout();

                    if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                    {
                        if (ViewboxFloatingBar.Margin == new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200))
                        {
                            ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);
                        }
                    }
                    else if (Settings.Appearance.IsFloatBarShowOnRight)
                    {
                        ViewboxFloatingBar.Margin = new Thickness(ViewboxFloatingBar.Margin.Left + width - ViewboxFloatingBar.ActualWidth, ViewboxFloatingBar.Margin.Top, ViewboxFloatingBar.Margin.Right, ViewboxFloatingBar.Margin.Bottom);
                    }
                }
            }

            SetColors();
        }

        private void SymbolIconDelete_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender != lastBorderMouseDownObject) return;
            if (inkCanvas.GetSelectedStrokes().Count > 0)
            {
                inkCanvas.Strokes.Remove(inkCanvas.GetSelectedStrokes());
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
            }
            else if (inkCanvas.Strokes.Count > 0)
            {
                if (Settings.Automation.IsAutoSaveStrokesAtClear && inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber)
                {
                    if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                        SaveScreenShot(true, $"{pptName}/{previousSlideID}_{DateTime.Now:HH-mm-ss}");
                    else
                        SaveScreenShot(true);
                }
                BtnClear_Click(BtnClear, null);
            }
            else
            {
                if (currentMode == 0 && BtnPPTSlideShowEnd.Visibility != Visibility.Visible)
                {
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                }
            }
        }

        private void SymbolIconSettings_Click(object sender, RoutedEventArgs e)
        {
            BtnSettings_Click(BtnSettings, null);
            HideSubPanels();
        }

        private void SymbolIconSelect_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnSelect_Click(BtnSelect, null);

            ImageEraser.Visibility = Visibility.Visible;
            ViewboxBtnColorBlackContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorBlueContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorGreenContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorRedContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorYellowContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorWhiteContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));

            HideSubPanels();
        }

        private void SymbolIconScreenshot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnScreenshot_Click(BtnScreenshot, null);
        }

        Point pointDesktop = new Point(-1, -1); //用于记录上次进入PPT或白板时的坐标
        Point pointPPT = new Point(-1, -1); //用于记录上次在PPT中打开白板时的坐标

        private async void ImageBlackboard_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (currentMode == 0)
            {
                //进入黑板
                if (BtnPPTSlideShowEnd.Visibility == Visibility.Collapsed)
                {
                    if (Main_Grid.Background != Brushes.Transparent)
                    {
                        await Dispatcher.BeginInvoke(new Action(() =>
                        {
                            SymbolIconCursor_Click(null, null);
                        }), DispatcherPriority.Loaded);
                    }
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        pointDesktop = new Point(ViewboxFloatingBar.Margin.Left, ViewboxFloatingBar.Margin.Top);
                    }), DispatcherPriority.Render);
                }
                else
                {
                    pointPPT = new Point(ViewboxFloatingBar.Margin.Left, ViewboxFloatingBar.Margin.Top);
                    lockSmithPPT = _lockSmith;
                    ChangeLockSmithState(lockSmithDesktop);
                }
                if (Settings.Canvas.UsingWhiteboard)
                {
                    BorderPenColorBlack_MouseUp(BorderPenColorBlack, null);
                }
                else
                {
                    BorderPenColorWhite_MouseUp(BorderPenColorWhite, null);
                }
            }
            else
            {
                //关闭黑板
                if (isInMultiTouchMode) BorderMultiTouchMode_MouseUp(null, null);

                if (BtnPPTSlideShowEnd.Visibility != Visibility.Collapsed)
                {
                    lockSmithDesktop = _lockSmith;
                    ChangeLockSmithState(lockSmithPPT);
                }
                BorderPenColorRed_MouseUp(BorderPenColorRed, null);
            }
            BtnSwitch_Click(BtnSwitch, null);
            SetColors();
            SetColorByIndex();

            if (currentMode != 0)
            {
                if (Settings.Canvas.UsingWhiteboard)
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        BorderPenColorBlack_MouseUp(BorderPenColorBlack, null);
                    }), DispatcherPriority.Loaded);
                }
                else
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        BorderPenColorWhite_MouseUp(BorderPenColorWhite, null);
                    }), DispatcherPriority.Loaded);
                }

                SetBorderFloatingBarMainControlsVisibility(true, false);

                //ViewboxFloatingBar.Margin = new Thickness(10, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);

                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);
                }), DispatcherPriority.Render);
            }
            else
            {
                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    BorderPenColorRed_MouseUp(BorderPenColorRed, null);
                }), DispatcherPriority.Loaded);

                if (BtnPPTSlideShowEnd.Visibility == Visibility.Collapsed)
                {
                    if (pointDesktop != new Point(-1, -1))
                    {
                        await Dispatcher.BeginInvoke(new Action(() =>
                        {
                            ViewboxFloatingBar.Margin = new Thickness(pointDesktop.X, pointDesktop.Y, -2000, -200);
                            pointDesktop = new Point(-1, -1);
                            if (Settings.Appearance.IsAutoCollapseFloatBar)
                            {
                                SetBorderFloatingBarMainControlsVisibility(false);
                            }
                        }), DispatcherPriority.Loaded);
                    }
                }
                else
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);
                    }), DispatcherPriority.Render);
                }
            }

            BtnExit.Foreground = Brushes.White;
            ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
            if (currentMode == 0 && inkCanvas.Strokes.Count == 0 && BtnPPTSlideShowEnd.Visibility != Visibility.Visible)
            {
                BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
            }
        }

        private void ImageEraser_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnErase_Click(BtnErase, e);

            ViewboxBtnColorBlackContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorBlueContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorGreenContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorRedContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorYellowContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorWhiteContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));

            HideSubPanels();
        }

        private void ImageCountdownTimer_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BtnCountdownTimer_Click(BtnCountdownTimer, null);
        }

        private void SymbolIconRand_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BtnRand_Click(BtnRand, null);
        }

        private void SymbolIconRandOne_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            new RandWindow(true).ShowDialog();
        }

        private void GridInkReplayButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BorderDrawShape.Visibility = Visibility.Collapsed;

            InkCanvasForInkReplay.Visibility = Visibility.Visible;
            inkCanvas.Visibility = Visibility.Collapsed;
            isStopInkReplay = false;
            InkCanvasForInkReplay.Strokes.Clear();
            StrokeCollection strokes = inkCanvas.Strokes.Clone();
            if (inkCanvas.GetSelectedStrokes().Count != 0)
            {
                strokes = inkCanvas.GetSelectedStrokes().Clone();
            }
            int k = 1, i = 0;
            new Thread(new ThreadStart(() =>
            {
                foreach (Stroke stroke in strokes)
                {
                    //Thread.Sleep(100);
                    //Application.Current.Dispatcher.Invoke(() =>
                    //{
                    //    InkCanvasForInkReplay.Strokes.Add(stroke);
                    //});
                    StylusPointCollection stylusPoints = new StylusPointCollection();
                    if (stroke.StylusPoints.Count == 629) //圆或椭圆
                    {
                        Stroke s = null;
                        foreach (StylusPoint stylusPoint in stroke.StylusPoints)
                        {
                            if (i++ >= 50)
                            {
                                i = 0;
                                Thread.Sleep(10);
                                if (isStopInkReplay) return;
                            }
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    InkCanvasForInkReplay.Strokes.Remove(s);
                                }
                                catch { }
                                stylusPoints.Add(stylusPoint);
                                s = new Stroke(stylusPoints.Clone());
                                s.DrawingAttributes = stroke.DrawingAttributes;
                                InkCanvasForInkReplay.Strokes.Add(s);
                            });
                        }
                    }
                    else
                    {
                        Stroke s = null;
                        foreach (StylusPoint stylusPoint in stroke.StylusPoints)
                        {
                            if (i++ >= k)
                            {
                                i = 0;
                                Thread.Sleep(10);
                                if (isStopInkReplay) return;
                            }
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    InkCanvasForInkReplay.Strokes.Remove(s);
                                }
                                catch { }
                                stylusPoints.Add(stylusPoint);
                                s = new Stroke(stylusPoints.Clone());
                                s.DrawingAttributes = stroke.DrawingAttributes;
                                InkCanvasForInkReplay.Strokes.Add(s);
                            });
                        }
                    }
                }
                Thread.Sleep(100);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    InkCanvasForInkReplay.Visibility = Visibility.Collapsed;
                    inkCanvas.Visibility = Visibility.Visible;
                });
            })).Start();
        }
        bool isStopInkReplay = false;
        private void InkCanvasForInkReplay_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                InkCanvasForInkReplay.Visibility = Visibility.Collapsed;
                inkCanvas.Visibility = Visibility.Visible;
                isStopInkReplay = true;
            }
        }

        private void SymbolIconTools_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (BorderTools.Visibility == Visibility.Visible)
            {
                BorderTools.Visibility = Visibility.Collapsed;
            }
            else
            {
                BorderTools.Visibility = Visibility.Visible;
            }
        }


        #region Drag

        bool isDragDropInEffect = false;
        Point pos = new Point();
        Point downPos = new Point();

        void Element_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragDropInEffect)
            {
                FrameworkElement currEle = sender as FrameworkElement;
                double xPos = e.GetPosition(null).X - pos.X + currEle.Margin.Left;
                double yPos = e.GetPosition(null).Y - pos.Y + currEle.Margin.Top;
                currEle.Margin = new Thickness(xPos, yPos, 0, 0);
                pos = e.GetPosition(null);
            }
        }

        void Element_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {

            FrameworkElement fEle = sender as FrameworkElement;
            isDragDropInEffect = true;
            pos = e.GetPosition(null);
            fEle.CaptureMouse();
            fEle.Cursor = Cursors.Hand;
        }

        void Element_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (isDragDropInEffect)
            {
                FrameworkElement ele = sender as FrameworkElement;
                isDragDropInEffect = false;
                ele.ReleaseMouseCapture();
            }
        }


        void SymbolIconEmoji_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragDropInEffect)
            {
                double xPos = e.GetPosition(null).X - pos.X + ViewboxFloatingBar.Margin.Left;
                double yPos = e.GetPosition(null).Y - pos.Y + ViewboxFloatingBar.Margin.Top;
                ViewboxFloatingBar.Margin = new Thickness(xPos, yPos, -2000, -200);
                pos = e.GetPosition(null);
            }
        }

        void SymbolIconEmoji_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isDragDropInEffect = true;
            pos = e.GetPosition(null);
            downPos = e.GetPosition(null);
            GridForFloatingBarDraging.Visibility = Visibility.Visible;

            SymbolIconEmoji.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Emoji;
        }

        void SymbolIconEmoji_MouseUp(object sender, MouseButtonEventArgs e)
        {
            isDragDropInEffect = false;

            if (e is null || (downPos.X == e.GetPosition(null).X && downPos.Y == e.GetPosition(null).Y))
            {
                SetBorderFloatingBarMainControlsVisibility(!borderFloatingBarMainControlsVisibility);
            }
            else if (Settings.Appearance.IsRememberFloatBarPosition)
            {
                SaveFloatBarPositionToFile();
            }

            GridForFloatingBarDraging.Visibility = Visibility.Collapsed;
            SymbolIconEmoji.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Emoji2;
        }

        bool borderFloatingBarMainControlsVisibility = true;
        void SetBorderFloatingBarMainControlsVisibility(bool isVisible, bool isAnimated = true)
        {
            borderFloatingBarMainControlsVisibility = isVisible;
            if (!isVisible)
            {
                BorderFloatingBarMainControls.RenderTransform.BeginAnimation(ScaleTransform.ScaleXProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(isAnimated ? 100 : 0))
                {
                    EasingFunction = new PowerEase() { Power = 4, EasingMode = EasingMode.EaseOut },
                });
                BorderFloatingBarMainControls.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(isAnimated ? 100 : 0)));
            }
            else
            {
                BorderFloatingBarMainControls.RenderTransform.BeginAnimation(ScaleTransform.ScaleXProperty, new DoubleAnimation(1, TimeSpan.FromMilliseconds(isAnimated ? 160 : 0))
                {
                    EasingFunction = new PowerEase() { Power = 4, EasingMode = EasingMode.EaseOut },
                });
                BorderFloatingBarMainControls.BeginAnimation(OpacityProperty, new DoubleAnimation(1, TimeSpan.FromMilliseconds(isAnimated ? 160 : 0)));
            }
        }

        #endregion


        private void GridPPTControlPrevious_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnPPTSlidesUp_Click(BtnPPTSlidesUp, null);
        }

        private void GridPPTControlNext_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnPPTSlidesDown_Click(BtnPPTSlidesDown, null);
        }

        private void ImagePPTControlEnd_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnPPTSlideShowEnd_Click(BtnPPTSlideShowEnd, null);
        }

        #endregion

        #region Save & Open

        private void SymbolIconSaveStrokes_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender || inkCanvas.Visibility != Visibility.Visible) return;

            BorderTools.Visibility = Visibility.Collapsed;

            GridNotifications.Visibility = Visibility.Collapsed;

            SaveInkCanvasStrokes();
        }

        private void SaveInkCanvasStrokes(bool newNotice = true)
        {
            try
            {
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\User Saved"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\User Saved");
                }

                FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                    @"\Ink Canvas Strokes\User Saved\" + DateTime.Now.ToString("u").Replace(':', '-') + ".icstk", FileMode.Create); //Ink Canvas STroKes
                inkCanvas.Strokes.Save(fs);

                if (newNotice)
                {
                    ShowNotification("墨迹成功保存至 " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                        @"\Ink Canvas Strokes\User Saved\" + DateTime.Now.ToString("u").Replace(':', '-') + ".icstk");
                }
                else
                {
                    AppendNotification("墨迹成功保存至 " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                        @"\Ink Canvas Strokes\User Saved\" + DateTime.Now.ToString("u").Replace(':', '-') + ".icstk");
                }
            }
            catch (Exception ex)
            {
                ShowNotification($"墨迹保存失败：{ex.Message}");
            }
        }

        private void SymbolIconPin_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ChangeLockSmithState(!_lockSmith);
        }

        private void SymbolIconOpenStrokes_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            BorderTools.Visibility = Visibility.Collapsed;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            string defaultFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\User Saved";
            if (Directory.Exists(defaultFolderPath))
            {
                openFileDialog.InitialDirectory = defaultFolderPath;
            }
            openFileDialog.Title = "打开墨迹文件";
            openFileDialog.Filter = "Ink Canvas Strokes File (*.icstk)|*.icstk";
            if (openFileDialog.ShowDialog() == true)
            {
                LogHelper.WriteLogToFile(string.Format("Strokes Insert: Name: {0}", openFileDialog.FileName), LogHelper.LogType.Event);
                try
                {
                    var fileStreamHasNoStroke = false;
                    using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        var strokes = new StrokeCollection(fs);
                        fileStreamHasNoStroke = strokes.Count == 0;
                        if (!fileStreamHasNoStroke)
                        {
                            ClearStrokes(true);
                            timeMachine.ClearStrokeHistory();
                            inkCanvas.Strokes.Add(strokes);
                            LogHelper.NewLog(string.Format("Strokes Insert: Strokes Count: {0}", inkCanvas.Strokes.Count.ToString()));
                        }
                    }
                    if (fileStreamHasNoStroke)
                    {
                        using (var ms = new MemoryStream(File.ReadAllBytes(openFileDialog.FileName)))
                        {
                            ms.Seek(0, SeekOrigin.Begin);
                            var strokes = new StrokeCollection(ms);
                            ClearStrokes(true);
                            timeMachine.ClearStrokeHistory();
                            inkCanvas.Strokes.Add(strokes);
                            LogHelper.NewLog(string.Format("Strokes Insert (2): Strokes Count: {0}", strokes.Count.ToString()));
                        }
                    }

                    if (inkCanvas.Visibility != Visibility.Visible)
                    {
                        SymbolIconCursor_Click(sender, null);
                    }
                }
                catch
                {
                    ShowNotification("墨迹打开失败");
                }
            }
        }



        #endregion

        #region Multi-finger Inking


        #endregion

    }
}
