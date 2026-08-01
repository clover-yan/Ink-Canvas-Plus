using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace InkCanvasPlus
{
    /// <summary>
    /// WelcomeWindow.xaml 的交互逻辑
    /// </summary>
    public partial class WelcomeWindow : Window
    {
        public WelcomeWindow()
        {
            InitializeComponent();
            TextBlockVersion.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            TextBlockUpdateContent.Inlines.Add(new Bold(new Run("【新功能】")));
            TextBlockUpdateContent.Inlines.Add(@"
本次更新增加了很多新设置项，点击工具栏中的「齿轮」按钮可以设置这些功能。

· 新增「荧光笔」功能，使用工具栏中的荧光笔图标切换，快捷键 Alt+M
· 工具栏图标支持缩放，默认大小为 80%，可按喜好调整显示大小
  （外观→工具栏图标缩放）
· 新增支持「垂直排布的 PPT 导航按钮」，适应不同的 PPT 课件
  （外观→显示垂直 PPT 导航按钮）
· 浮动工具栏可以「显示在屏幕右侧」并向左展开，适应不同的操作环境
  （外观→浮动栏显示在右侧）
· 浮动工具栏可以「记住上次显示的位置」，下次启动自动还原
  （外观→记忆浮动栏位置）
· 橡皮擦增加「极小」「很小」两档尺寸，并优化了大小计算逻辑
  （画板→橡皮大小）
· 新增「套索选择」快捷键 Alt+Q，用键盘操作更顺手
  提示：您可以在设置界面的底部查看完整的快捷键列表。

");
            TextBlockUpdateContent.Inlines.Add(new Bold(new Run("【改进】")));
            TextBlockUpdateContent.Inlines.Add(@"
· 优化多实例处理，若 Ink Canvas Plus 已在运行但无响应，可以重新启动它
· 手动检查更新时若没有检查到更新，现在会给出提示
· 更新提示窗口不会再让主窗口卡死
· 下载与联系方式等链接已迁移至新官网 cloveryan.com

");
            TextBlockUpdateContent.Inlines.Add(new Bold(new Run("【问题修复】")));
            TextBlockUpdateContent.Inlines.Add(@"
· 修复启动时浮动工具栏自动折叠不生效的问题
· 修复使用鼠标和手写笔时无法拖拽或操作选区手柄的问题
· 修复切换鼠标与画笔模式时偶尔界面闪动的问题
· 修复切换回画笔模式后选定的颜色被重置的问题
· 修复进入和退出黑板时画笔颜色初始化不一致的问题
· 修复窗口可能被意外移动或改变大小的问题

");
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ButtonShowMore_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://cloveryan.com/apps/Ink-Canvas-Plus/changelog");
        }

        private void HyperlinkButtonWebsite_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://cloveryan.com/ic+");
        }
    }
}
