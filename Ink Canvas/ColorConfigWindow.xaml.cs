using InkCanvasPlus.Helpers;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using MessageBox = System.Windows.MessageBox;

namespace InkCanvasPlus
{
    /// <summary>
    /// ColorConfigWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ColorConfigWindow : Window
    {
        private Border[] lightBorders;
        private Border[] darkBorders;
        private TextBox[] lightTextBoxes;
        private TextBox[] darkTextBoxes;
        private Color[] lightColors;
        private Color[] darkColors;
        private bool isSaved = true;

        public ColorConfigWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            lightBorders = new[] { BorderLightRed, BorderLightGreen, BorderLightBlue, BorderLightYellow };
            darkBorders = new[] { BorderDarkRed, BorderDarkGreen, BorderDarkBlue, BorderDarkYellow };
            lightTextBoxes = new[] { TextBoxLightRed, TextBoxLightGreen, TextBoxLightBlue, TextBoxLightYellow };
            darkTextBoxes = new[] { TextBoxDarkRed, TextBoxDarkGreen, TextBoxDarkBlue, TextBoxDarkYellow };

            lightColors = ColorConfigHelper.LoadColors(ColorConfigHelper.LightColorFile, ColorConfigHelper.DefaultLightColors);
            darkColors = ColorConfigHelper.LoadColors(ColorConfigHelper.DarkColorFile, ColorConfigHelper.DefaultDarkColors);

            RefreshUI();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!isSaved)
            {
                var result = MessageBox.Show("有未保存的颜色更改，是否保存？", "配置画笔颜色", MessageBoxButton.YesNoCancel);
                if (result == MessageBoxResult.Yes)
                {
                    SaveColors();
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                }
            }
        }

        private void RefreshUI()
        {
            for (int i = 0; i < 4; i++)
            {
                SetColorUI(i, true, lightColors[i]);
                SetColorUI(i, false, darkColors[i]);
            }
            isSaved = true;
        }

        private void SetColorUI(int index, bool isLight, Color color)
        {
            var borders = isLight ? lightBorders : darkBorders;
            var textBoxes = isLight ? lightTextBoxes : darkTextBoxes;
            borders[index].Background = new SolidColorBrush(color);
            textBoxes[index].Text = ColorConfigHelper.ColorToArgbString(color);
        }

        private void SetColor(int index, bool isLight, Color color)
        {
            if (isLight) lightColors[index] = color;
            else darkColors[index] = color;
            SetColorUI(index, isLight, color);
            isSaved = false;
        }

        private bool TryGetSlot(object tag, out bool isLight, out int index)
        {
            isLight = false;
            index = -1;
            string s = tag as string;
            if (s == null || s.Length != 2) return false;
            if (s[0] == 'L') isLight = true;
            else if (s[0] == 'D') isLight = false;
            else return false;
            index = s[1] - '0';
            return index >= 0 && index < 4;
        }

        private void ColorSwatch_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (TryGetSlot((sender as FrameworkElement)?.Tag, out bool isLight, out int index))
            {
                PickColor(index, isLight);
            }
        }

        private void PickColor(int index, bool isLight)
        {
            var current = isLight ? lightColors[index] : darkColors[index];
            using (var dialog = new System.Windows.Forms.ColorDialog())
            {
                dialog.FullOpen = true;
                dialog.Color = System.Drawing.Color.FromArgb(current.A, current.R, current.G, current.B);
                if (ShowColorDialog(dialog))
                {
                    var c = dialog.Color;
                    SetColor(index, isLight, Color.FromArgb(c.A, c.R, c.G, c.B));
                }
            }
        }

        private bool ShowColorDialog(System.Windows.Forms.ColorDialog dialog)
        {
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero) helper.EnsureHandle();
            var nativeWindow = new System.Windows.Forms.NativeWindow();
            nativeWindow.AssignHandle(helper.Handle);
            try
            {
                return dialog.ShowDialog(nativeWindow) == System.Windows.Forms.DialogResult.OK;
            }
            finally
            {
                nativeWindow.ReleaseHandle();
            }
        }

        private void HexTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox == null || lightColors == null) return;
            if (!TryGetSlot(textBox.Tag, out bool isLight, out int index)) return;
            try
            {
                Color color = ColorConfigHelper.ArgbStringToColor(textBox.Text.Trim());
                SetColor(index, isLight, color);
            }
            catch
            {
                textBox.Text = ColorConfigHelper.ColorToArgbString(isLight ? lightColors[index] : darkColors[index]);
            }
        }

        private void BtnResetDefault_Click(object sender, RoutedEventArgs e)
        {
            lightColors = (Color[])ColorConfigHelper.DefaultLightColors.Clone();
            darkColors = (Color[])ColorConfigHelper.DefaultDarkColors.Clone();
            RefreshUI();
            isSaved = false;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveColors();
            isSaved = true;
            ShowSavedIndicator();
        }

        private int savedIndicatorToken = 0;

        private void ShowSavedIndicator()
        {
            int token = ++savedIndicatorToken;
            BorderSavedIndicator.BeginAnimation(OpacityProperty, null);
            BorderSavedIndicator.Opacity = 0;

            var storyboard = new Storyboard();

            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(150));
            Storyboard.SetTarget(fadeIn, BorderSavedIndicator);
            Storyboard.SetTargetProperty(fadeIn, new PropertyPath(OpacityProperty));
            storyboard.Children.Add(fadeIn);

            var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(200))
            {
                BeginTime = TimeSpan.FromSeconds(2)
            };
            Storyboard.SetTarget(fadeOut, BorderSavedIndicator);
            Storyboard.SetTargetProperty(fadeOut, new PropertyPath(OpacityProperty));
            storyboard.Children.Add(fadeOut);

            storyboard.Completed += (s, e) =>
            {
                if (token != savedIndicatorToken) return;
                BorderSavedIndicator.Opacity = 0;
            };
            storyboard.Begin();
        }

        private void SaveColors()
        {
            ColorConfigHelper.SaveColors(ColorConfigHelper.LightColorFile, lightColors);
            ColorConfigHelper.SaveColors(ColorConfigHelper.DarkColorFile, darkColors);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
