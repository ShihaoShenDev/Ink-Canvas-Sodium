using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Ink_Canvas.Popups;
using Ink_Canvas.Windows.SettingsViews;
using iNKORE.UI.WPF.Helpers;
using iNKORE.UI.WPF.Modern.Controls;
using OSVersionExtension;
using Ink_Canvas.Helpers;

namespace Ink_Canvas.Windows {
    public partial class SettingsWindow : Window {

        public SettingsWindow() {
            try {
                InitializeComponent();

                // 初始化侧边栏项目
                SidebarItemsControl.ItemsSource = SidebarItems;
                
                // 安全地获取图标资源
                DrawingImage GetIconResource(string key) {
                    try {
                        return FindResource(key) as DrawingImage;
                    }
                    catch (Exception) {
                        return null;
                    }
                }
                
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "启动时行为",
                    Name = "StartupItem",
                    IconSource = GetIconResource("StartupIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "画板和墨迹",
                    Name = "CanvasAndInkItem",
                    IconSource = GetIconResource("CanvasAndInkIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "手势操作",
                    Name = "GesturesItem",
                    IconSource = GetIconResource("GesturesIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Separator
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "个性化和外观",
                    Name = "AppearanceItem",
                    IconSource = GetIconResource("AppearanceIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "墨迹转形状",
                    Name = "InkRecognitionItem",
                    IconSource = GetIconResource("InkRecognitionIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "几何与形状绘制",
                    Name = "ShapeDrawingItem",
                    IconSource = GetIconResource("ShapeDrawingIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "自动化行为",
                    Name = "AutomationItem",
                    IconSource = GetIconResource("AutomationIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Separator
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "PowerPoint 支持",
                    Name = "PowerPointItem",
                    IconSource = GetIconResource("PowerPointIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "插件和脚本",
                    Name = "ExtensionsItem",
                    IconSource = GetIconResource("ExtensionsIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Separator
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "存储空间",
                    Name = "StorageItem",
                    IconSource = GetIconResource("StorageIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "截图和屏幕捕捉",
                    Name = "SnapshotItem",
                    IconSource = GetIconResource("SnapshotIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "点名器设置",
                    Name = "LuckyRandomItem",
                    IconSource = GetIconResource("LuckyRandomIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "高级选项",
                    Name = "AdvancedItem",
                    IconSource = GetIconResource("AdvancedIcon"),
                    Selected = false,
                });
                SidebarItems.Add(new SidebarItem() {
                    Type = SidebarItemType.Item,
                    Title = "关于 InkCanvasForClass",
                    Name = "AboutItem",
                    IconSource = GetIconResource("AboutIcon"),
                    Selected = false,
                });
                _selectedSidebarItemName = "AboutItem";
                UpdateSidebarItemsSelection();

                // 安全地初始化数组，检查每个元素是否为null
                var settingsPanesList = new List<Grid>();
                if (AboutPane != null) settingsPanesList.Add(AboutPane);
                if (ExtensionsPane != null) settingsPanesList.Add(ExtensionsPane);
                if (CanvasAndInkPane != null) settingsPanesList.Add(CanvasAndInkPane);
                if (GesturesPane != null) settingsPanesList.Add(GesturesPane);
                if (StartupPane != null) settingsPanesList.Add(StartupPane);
                if (AppearancePane != null) settingsPanesList.Add(AppearancePane);
                if (InkRecognitionPane != null) settingsPanesList.Add(InkRecognitionPane);
                if (AutomationPane != null) settingsPanesList.Add(AutomationPane);
                if (PowerPointPane != null) settingsPanesList.Add(PowerPointPane);
                SettingsPanes = settingsPanesList.ToArray();

                var scrollViewersList = new List<ScrollViewer>();
                if (SettingsAboutPanel?.AboutScrollViewerEx != null) 
                    scrollViewersList.Add(SettingsAboutPanel.AboutScrollViewerEx);
                if (CanvasAndInkScrollViewerEx != null) 
                    scrollViewersList.Add(CanvasAndInkScrollViewerEx);
                if (GesturesScrollViewerEx != null) 
                    scrollViewersList.Add(GesturesScrollViewerEx);
                if (StartupScrollViewerEx != null) 
                    scrollViewersList.Add(StartupScrollViewerEx);
                if (AppearancePane?.Children.Count > 0 && AppearancePane.Children[0] is AppearancePanel appearancePanel && 
                    appearancePanel.BaseView?.SettingsViewScrollViewer != null)
                    scrollViewersList.Add(appearancePanel.BaseView.SettingsViewScrollViewer);
                if (InkRecognitionScrollViewerEx != null) 
                    scrollViewersList.Add(InkRecognitionScrollViewerEx);
                if (AutomationScrollViewerEx != null) 
                    scrollViewersList.Add(AutomationScrollViewerEx);
                if (PowerPointScrollViewerEx != null) 
                    scrollViewersList.Add(PowerPointScrollViewerEx);
                SettingsPaneScrollViewers = scrollViewersList.ToArray();

                if (SettingsAboutPanel != null) {
                    SettingsAboutPanel.IsTopBarNeedShadowEffect += (o, s) => {
                        if (DropShadowEffectTopBar != null) DropShadowEffectTopBar.Opacity = 0.25;
                    };
                    SettingsAboutPanel.IsTopBarNeedNoShadowEffect += (o, s) => {
                        if (DropShadowEffectTopBar != null) DropShadowEffectTopBar.Opacity = 0;
                    };
                }
            }
            catch (Exception ex) {
                LogHelper.WriteLogToFile($"Error initializing SettingsWindow: {ex}", LogHelper.LogType.Error);
                System.Windows.MessageBox.Show($"设置窗口初始化失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public Grid[] SettingsPanes = new Grid[] { };
        public ScrollViewer[] SettingsPaneScrollViewers = new ScrollViewer[] { };

        public enum SidebarItemType {
            Item,
            Separator
        }

        public class SidebarItem {
            public SidebarItemType Type { get; set; }
            public string Title { get; set; }
            public string Name { get; set; }
            public ImageSource IconSource { get; set; }
            public bool Selected { get; set; }
            public Visibility _spVisibility {
                get => this.Type == SidebarItemType.Separator ? Visibility.Visible : Visibility.Collapsed;
            }
            public Visibility _siVisibility {
                get => this.Type == SidebarItemType.Item ? Visibility.Visible : Visibility.Collapsed;
            }

            public SolidColorBrush _siBackground {
                get => this.Selected
                    ? new SolidColorBrush(Color.FromRgb(217, 217, 217))
                    : new SolidColorBrush(Colors.Transparent);
            }
        }

        public string _selectedSidebarItemName = "";
        public ObservableCollection<SidebarItem> SidebarItems = new ObservableCollection<SidebarItem>();

        public void UpdateSidebarItemsSelection() {
            try {
                foreach (var si in SidebarItems) {
                    si.Selected = si.Name == _selectedSidebarItemName;
                    if (si.Selected && SettingsWindowTitle != null) {
                        SettingsWindowTitle.Text = si.Title;
                    }
                }
                
                if (SidebarItems != null) {
                    CollectionViewSource.GetDefaultView(SidebarItems).Refresh();
                }

                // 安全地设置面板可见性
                if (AboutPane != null) AboutPane.Visibility = _selectedSidebarItemName == "AboutItem" ? Visibility.Visible : Visibility.Collapsed;
                if (ExtensionsPane != null) ExtensionsPane.Visibility = _selectedSidebarItemName == "ExtensionsItem" ? Visibility.Visible : Visibility.Collapsed;
                if (CanvasAndInkPane != null) CanvasAndInkPane.Visibility = _selectedSidebarItemName == "CanvasAndInkItem" ? Visibility.Visible : Visibility.Collapsed;
                if (GesturesPane != null) GesturesPane.Visibility = _selectedSidebarItemName == "GesturesItem" ? Visibility.Visible : Visibility.Collapsed;
                if (StartupPane != null) StartupPane.Visibility = _selectedSidebarItemName == "StartupItem" ? Visibility.Visible : Visibility.Collapsed;
                if (AppearancePane != null) AppearancePane.Visibility = _selectedSidebarItemName == "AppearanceItem" ? Visibility.Visible : Visibility.Collapsed;
                if (InkRecognitionPane != null) InkRecognitionPane.Visibility = _selectedSidebarItemName == "InkRecognitionItem" ? Visibility.Visible : Visibility.Collapsed;
                if (AutomationPane != null) AutomationPane.Visibility = _selectedSidebarItemName == "AutomationItem" ? Visibility.Visible : Visibility.Collapsed;
                if (PowerPointPane != null) PowerPointPane.Visibility = _selectedSidebarItemName == "PowerPointItem" ? Visibility.Visible : Visibility.Collapsed;
                
                // 安全地滚动所有 ScrollViewers
                if (SettingsPaneScrollViewers != null) {
                    foreach (var sv in SettingsPaneScrollViewers) {
                        if (sv != null) {
                            sv.ScrollToTop();
                        }
                    }
                }
            }
            catch (Exception ex) {
                LogHelper.WriteLogToFile($"Error in UpdateSidebarItemsSelection: {ex}", LogHelper.LogType.Error);
            }
        }

        private void ScrollViewerEx_ScrollChanged(object sender, ScrollChangedEventArgs e) {
            try {
                var scrollViewer = sender as ScrollViewer;
                if (scrollViewer == null) return;
                
                if (scrollViewer.VerticalOffset >= 10) {
                    if (DropShadowEffectTopBar != null) DropShadowEffectTopBar.Opacity = 0.25;
                } else {
                    if (DropShadowEffectTopBar != null) DropShadowEffectTopBar.Opacity = 0;
                }
            }
            catch (Exception ex) {
                LogHelper.WriteLogToFile($"Error in ScrollViewerEx_ScrollChanged: {ex}", LogHelper.LogType.Error);
            }
        }

        private void ScrollBar_Scroll(object sender, RoutedEventArgs e) {
            var scrollbar = (ScrollBar)sender;
            var scrollviewer = scrollbar.FindAscendant<ScrollViewer>();
            if (scrollviewer != null) scrollviewer.ScrollToVerticalOffset(scrollbar.Track.Value);
        }

        private void ScrollBarTrack_MouseEnter(object sender, MouseEventArgs e) {
            var border = (Border)sender;
            if (border.Child is Track track) {
                track.Width = 16;
                track.Margin = new Thickness(0, 0, -2, 0);
                var scrollbar = track.FindAscendant<ScrollBar>();
                if (scrollbar != null) scrollbar.Width = 16;
                var grid = track.FindAscendant<Grid>();
                if (grid.FindDescendantByName("ScrollBarBorderTrackBackground") is Border backgroundBorder) {
                    backgroundBorder.Width = 8;
                    backgroundBorder.CornerRadius = new CornerRadius(4);
                    backgroundBorder.Opacity = 1;
                }
                var thumb = track.Thumb.Template.FindName("ScrollbarThumbEx", track.Thumb) ;
                if (thumb != null) {
                    var _thumb = thumb as Border;
                    _thumb.CornerRadius = new CornerRadius(4);
                    _thumb.Width = 8;
                    _thumb.Margin = new Thickness(-0.75, 0, 1, 0);
                    _thumb.Background = new SolidColorBrush(Color.FromRgb(138, 138, 138));
                }
            }
        }

        private void ScrollBarTrack_MouseLeave(object sender, MouseEventArgs e) {
            var border = (Border)sender;
            border.Background = new SolidColorBrush(Colors.Transparent);
            border.CornerRadius = new CornerRadius(0);
            if (border.Child is Track track) {
                track.Width = 6;
                track.Margin = new Thickness(0, 0, 0, 0);
                var scrollbar = track.FindAscendant<ScrollBar>();
                if (scrollbar != null) scrollbar.Width = 6;
                var grid = track.FindAscendant<Grid>();
                if (grid.FindDescendantByName("ScrollBarBorderTrackBackground") is Border backgroundBorder) {
                    backgroundBorder.Width = 3;
                    backgroundBorder.CornerRadius = new CornerRadius(1.5);
                    backgroundBorder.Opacity = 0;
                }
                var thumb = track.Thumb.Template.FindName("ScrollbarThumbEx", track.Thumb) ;
                if (thumb != null) {
                    var _thumb = thumb as Border;
                    _thumb.CornerRadius = new CornerRadius(1.5);
                    _thumb.Width = 3;
                    _thumb.Margin = new Thickness(0);
                    _thumb.Background = new SolidColorBrush(Color.FromRgb(195, 195, 195));
                }
            }
        }

        private void ScrollbarThumb_MouseDown(object sender, MouseButtonEventArgs e) {
            var thumb = (Thumb)sender;
            var border = thumb.Template.FindName("ScrollbarThumbEx",thumb);
            ((Border)border).Background = new SolidColorBrush(Color.FromRgb(95, 95, 95));
        }

        private void ScrollbarThumb_MouseUp(object sender, MouseButtonEventArgs e) {
            var thumb = (Thumb)sender;
            var border = thumb.Template.FindName("ScrollbarThumbEx",thumb);
            ((Border)border).Background = new SolidColorBrush(Color.FromRgb(138, 138, 138));
        }

        private Border _sidebarItemMouseDownBorder = null;

        private void SidebarItem_MouseDown(object sender, MouseButtonEventArgs e) {
            if (_sidebarItemMouseDownBorder != null || _sidebarItemMouseDownBorder == sender) return;
            _sidebarItemMouseDownBorder = (Border)sender;
            var bd = sender as Border;
            if (bd.FindDescendantByName("MouseFeedbackBorder") is Border feedbackBd) feedbackBd.Opacity = 0.12;
        }

        private void SidebarItem_MouseUp(object sender, MouseButtonEventArgs e) {
            if (_sidebarItemMouseDownBorder == null || _sidebarItemMouseDownBorder != sender) return;
            if (_sidebarItemMouseDownBorder.Tag is SidebarItem data) _selectedSidebarItemName = data.Name;
            SidebarItem_MouseLeave(sender, null);
            UpdateSidebarItemsSelection();
        }

        private void SidebarItem_MouseLeave(object sender, MouseEventArgs e) {
            if (_sidebarItemMouseDownBorder == null || _sidebarItemMouseDownBorder != sender) return;
            if (_sidebarItemMouseDownBorder.FindDescendantByName("MouseFeedbackBorder") is Border feedbackBd) feedbackBd.Opacity = 0;
            _sidebarItemMouseDownBorder = null;
        }
    }
}
