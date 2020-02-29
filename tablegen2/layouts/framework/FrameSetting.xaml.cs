using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using tablegen2.logic;
using System.Collections.Generic;
namespace tablegen2.layouts
{
    /// <summary>
    /// FrameSetting.xaml 的交互逻辑
    /// </summary>
    public partial class FrameSetting : UserControl
    {
        public event Action ExcelDirChanged;
        public event Action ExportDirChanged;
        public event Action ExportFormatChanged;
        public event Action MoreSettingEvent;

        public FrameSetting()
        {
            InitializeComponent();
            if (AppData.Config != null)
            {
                cbExportFormat.SelectComboBoxItemByTag(AppData.Config.ExportFormat.ToString());
                indentCheckBox.IsChecked = AppData.Config.OutputLuaWithIndent;
                InitComboBox();
            }
        }

        public void InitComboBox()
        {
            if(AppData.Config.ExcelDirHistory != null)
            {
                boxExcelDir.Items.Clear();
                List<string> key1 = new List<string>(AppData.Config.ExcelDirHistory.Keys);
                for (int i = key1.Count - 1; i >= 0; i--)
                {
                    boxExcelDir.Items.Add(key1[i]);
                }
                boxExcelDir.SelectedItem = AppData.Config.ExcelDir;
                boxExcelDir.SelectionChanged += BoxExcelDir_SelectionChanged;

            }
  
            if(AppData.Config.ExportDirHistory != null)
            {
                boxExportDir.Items.Clear();
                List<string> key2 = new List<string>(AppData.Config.ExportDirHistory.Keys);
                for (int i = key2.Count - 1; i >= 0; i--)
                {
                    boxExportDir.Items.Add(key2[i]);
                }
                boxExportDir.SelectedItem = AppData.Config.ExportDir;
                boxExportDir.SelectionChanged += BoxExportDir_SelectionChanged;
            }

        }

        private void BoxExcelDir_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string dir = boxExcelDir.SelectedItem.ToString();
                AppData.Config.ExcelDir = dir;
                AppData.saveConfig();
            if (ExcelDirChanged != null)
                ExcelDirChanged.Invoke();
        }

        private void BoxExportDir_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string dir = boxExportDir.SelectedItem.ToString();
            AppData.Config.ExportDir = dir;
            AppData.saveConfig();
            if (ExportDirChanged != null)
                ExportDirChanged.Invoke();
        }

        public void browseExcelDirectory()
        {
            Util.performClick(btnBrowseExcelDir);
        }

        private void exportFormat_Selected(object sender, RoutedEventArgs e)
        {
            var item = sender as ComboBoxItem;
            var fmt = (TableExportFormat)Enum.Parse(typeof(TableExportFormat), item.Tag as string);
            if (fmt != TableExportFormat.Unknown)
            {
                AppData.Config.ExportFormat = fmt;
                AppData.saveConfig();
                if (ExportFormatChanged != null)
                    ExportFormatChanged.Invoke();
            }
        }


        private void btnBrowseExcelDir_Clicked(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择Excel目录";
            dialog.ShowNewFolderButton = true;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    //boxExcelDir.Text = dialog.SelectedPath;
                    AppData.Config.ExcelDir = dialog.SelectedPath;

                    if(AppData.Config.ExcelDirHistory == null)
                    {
                        AppData.Config.ExcelDirHistory = new Dictionary<string, int>();
                    }
                    if (AppData.Config.ExcelDirHistory.ContainsKey(dialog.SelectedPath))
                    {
                        AppData.Config.ExcelDirHistory[dialog.SelectedPath] += 1;
                    }
                    else
                    {
                        AppData.Config.ExcelDirHistory[dialog.SelectedPath] = 1;
                        boxExcelDir.Items.Add(dialog.SelectedPath);
                    }
                    boxExcelDir.SelectedItem = dialog.SelectedPath;
                    Dictionary<string, int> temp = AppData.Config.ExcelDirHistory.OrderBy(o => o.Value).ToDictionary(p=>p.Key, o=>o.Value);
                    AppData.Config.ExcelDirHistory = temp;

                    AppData.saveConfig();
                    if (ExcelDirChanged != null)
                        ExcelDirChanged.Invoke();
                }
            }
        }

        private void btnBrowseExportDir_Clicked(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择数据导出目录";
            dialog.ShowNewFolderButton = true;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    //txtExportDir.Text = dialog.SelectedPath;
                    AppData.Config.ExportDir = dialog.SelectedPath;

                    if (AppData.Config.ExportDirHistory == null)
                    {
                        AppData.Config.ExportDirHistory = new Dictionary<string, int>();
                    }
                    if (AppData.Config.ExportDirHistory.ContainsKey(dialog.SelectedPath))
                    {
                        AppData.Config.ExportDirHistory[dialog.SelectedPath] += 1;
                    }
                    else
                    {
                        AppData.Config.ExportDirHistory[dialog.SelectedPath] = 1;
                        boxExportDir.Items.Add(dialog.SelectedPath);
                        
                    }
                    boxExportDir.SelectedItem = dialog.SelectedPath;
                    Dictionary<string, int> temp = AppData.Config.ExportDirHistory.OrderBy(o => o.Value).ToDictionary(p => p.Key, o => o.Value);
                    AppData.Config.ExportDirHistory = temp;
                    AppData.saveConfig();
                    if (ExportDirChanged != null)
                        ExportDirChanged.Invoke();
                }
            }
        }

        private void btnOpenExcelDir_Clicked(object sender, RoutedEventArgs e)
        {
            var dir = AppData.Config.ExcelDir;
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                this.InfBox("请先使用‘浏览’功能选择合法的配置目录");
                return;
            }
            Util.OpenDir(dir);
        }

        private void btnOpenExportDir_Clicked(object sender, RoutedEventArgs e)
        {
            var dir = AppData.Config.ExportDir;
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                this.InfBox("请先使用‘浏览’功能选择合法的配置目录");
                return;
            }
            Util.OpenDir(dir);
        }

        private void btnMoreSetting_Clicked(object sender, RoutedEventArgs e)
        {
            if (MoreSettingEvent != null)
                MoreSettingEvent.Invoke();
        }

        private void indent_Checked(object sender, RoutedEventArgs e)
        {
            AppData.Config.OutputLuaWithIndent = true;
            AppData.saveConfig();
        }

        private void indent_Unchecked(object sender, RoutedEventArgs e)
        {
            AppData.Config.OutputLuaWithIndent = false;
            AppData.saveConfig();
        }

        private void fit_Unity3D(object sender, RoutedEventArgs e)
        {
            AppData.Config.FitUnity3D = true;
            AppData.saveConfig();
        }

        private void fit_Zendo(object sender, RoutedEventArgs e)
        {
            AppData.Config.FitUnity3D = false;
            AppData.saveConfig();
        }


    }
}
