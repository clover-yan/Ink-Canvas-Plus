using System;
using System.IO;
using System.Windows.Media;

namespace InkCanvasPlus.Helpers
{
    public static class ColorConfigHelper
    {
        public static readonly Color[] DefaultLightColors = new Color[]
        {
            (Color)ColorConverter.ConvertFromString("#FFFF3333"),
            (Color)ColorConverter.ConvertFromString("#FF1ED760"),
            (Color)ColorConverter.ConvertFromString("#FF239AD6"),
            (Color)ColorConverter.ConvertFromString("#FFFFC000"),
        };

        public static readonly Color[] DefaultDarkColors = new Color[]
        {
            (Color)ColorConverter.ConvertFromString("#FFFF0000"),
            (Color)ColorConverter.ConvertFromString("#FF169141"),
            (Color)ColorConverter.ConvertFromString("#FF239AD6"),
            (Color)ColorConverter.ConvertFromString("#FFF38B00"),
        };

        private static string ColorsFolder
        {
            get { return Path.Combine(App.RootPath, "Colors"); }
        }

        public static string LightColorFile
        {
            get { return Path.Combine(ColorsFolder, "Light.ini"); }
        }

        public static string DarkColorFile
        {
            get { return Path.Combine(ColorsFolder, "Dark.ini"); }
        }

        public static Color[] LoadColors(string colorFile, Color[] defaults)
        {
            try
            {
                if (File.Exists(colorFile))
                {
                    string[] lines = File.ReadAllLines(colorFile);
                    if (lines.Length >= 4)
                    {
                        Color[] colors = new Color[4];
                        for (int i = 0; i < 4; i++)
                        {
                            try { colors[i] = ArgbStringToColor(lines[i]); }
                            catch { colors[i] = defaults[i]; }
                        }
                        return colors;
                    }
                }
            }
            catch { }

            return (Color[])defaults.Clone();
        }

        public static void SaveColors(string colorFile, Color[] colors)
        {
            Directory.CreateDirectory(ColorsFolder);
            string[] lines = new string[4];
            for (int i = 0; i < 4; i++)
            {
                lines[i] = ColorToArgbString(colors[i]);
            }
            File.WriteAllLines(colorFile, lines);
        }

        public static string ColorToArgbString(Color color)
        {
            return string.Format("#{0:X2}{1:X2}{2:X2}{3:X2}", color.A, color.R, color.G, color.B);
        }

        public static Color ArgbStringToColor(string text)
        {
            return (Color)ColorConverter.ConvertFromString(text);
        }
    }
}
