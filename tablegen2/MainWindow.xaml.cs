using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using tablegen2.layouts;
using tablegen2.logic;

namespace tablegen2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            AppData.MainWindow = this;

            InitializeComponent();

            setting.ExcelDirChanged += () => tree.refreshExcelPath(AppData.Config.ExcelDir);
            setting.ExportFormatChanged += () => refreshButtonGenAll();
            setting.MoreSettingEvent += () => _flipMoreSettingPanel();
            tree.OpenExcelRequest += () => setting.browseExcelDirectory();

            if (AppData.Config != null)
                refreshButtonGenAll();
        }

        #region 事件处理
        private void Window_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                _flipHelpPanel();
                e.Handled = true;
            }
        }

        private void btnClear_Clicked(object sender, RoutedEventArgs e)
        {
            clearLog();
        }
        private void btnHelp_Clicked(object sender, RoutedEventArgs e)
        {
            _flipHelpPanel();
        }

        private void btnGenAll_Clicked(object sender, RoutedEventArgs e)
        {
            string excelDir = AppData.Config.ExcelDir;
            string exportDir = AppData.Config.ExportDir;
            TableExportFormat fmt = AppData.Config.ExportFormat;

            if (string.IsNullOrEmpty(excelDir) || !Directory.Exists(excelDir))
            {
                Log.Err("请选择合法的Excel配置目录！");
                return;
            }

            if (string.IsNullOrEmpty(exportDir) || !Directory.Exists(exportDir))
            {
                Log.Err("请选择合法的导出目录！");
                return;
            }

            if (fmt == TableExportFormat.Unknown)
            {
                Log.Err("请选择导出数据格式！");
                return;
            }

            var excels = tree.AllExcels.ToList();
            if (excels.Count == 0)
            {
                Log.Wrn("您选择的配置目录中不包含任何Excel文件！ 目录：{0}", excelDir);
                return;
            }

            Log.Msg("=================================================");
            int succ = 0, failed = 0;
            foreach (var filePath in excels)
            {
                bool succeed = _genSingleFileImpl(filePath, exportDir, fmt);

                if (succeed)
                    succ++;
                else
                    failed++;
            }

            if (failed > 0)
                Log.Err("========== {0} 个成功， {1} 个失败 ========== ", succ, failed);
            else
                Log.Suc("========== {0} 个成功 ========== ", succ);
        }
        #endregion

        public void addMessage(string msg, Color color)
        {
            if (console != null)
                console.addMessage(msg, color);
        }

        public void clearLog()
        {
            console.clearMessage();
        }

        public void refreshButtonGenAll()
        {
            switch (AppData.Config.ExportFormat)
            {
                case TableExportFormat.Unknown:
                    btnGenAll.IsEnabled = false;
                    btnGenAll.Content = "请选择生成的数据格式";
                    break;
                default:
                    btnGenAll.IsEnabled = true;
                    btnGenAll.Content = string.Format(btnGenAll.Tag as string, AppData.Config.ExportFormat.ToString());
                    break;
            }
        }

        public void editForExcel(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Log.Err("源文件不存在！ {0}", filePath);
                return;
            }

            try
            {
                TableExcelData data = TableExcelReader.loadFromExcel(filePath);
                openExcelView(data, filePath);
            }
            catch (System.Exception ex)
            {
                Log.Err(ex.Message);
            }
        }

        public void openExcelView(TableExcelData data, string filePath)
        {
            var panel = new FrameExcelView();
            panel.refreshUIByTableExcelData(data);
            panel.setFilePath(filePath);

            var pw = new PopupWindow(panel);
            pw.ResizeMode = ResizeMode.CanResize;
            pw.Owner = Window.GetWindow(this);
            pw.Title = string.Format("查看配置表 -- {0}", filePath);
            pw.MinWidth = 600;
            pw.MinHeight = 400;
            if (pw.ShowDialog() == true)
            {
            }
        }

        public void rectifyFileFormat(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Log.Err("源文件不存在！ {0}", filePath);
                return;
            }

            Log.Msg("=================================================");

            try
            {
                Log.Msg("正在优化 {0}", filePath);
                TableExcelData data = TableExcelReader.loadFromExcel(filePath);
                TableExcelWriter.genExcel(data, filePath);
                Log.Msg("优化完毕！");
            }
            catch (System.Exception ex)
            {
                Log.Err(ex.Message);
            }
        }

        public void rectifyAllFileFormat(List<string> files)
        {
            if (files.Count == 0)
            {
                Log.Wrn("没有要优化的文件");
                return;
            }

            Log.Msg("=================================================");
            foreach (var filePath in files)
            {
                try
                {
                    Log.Msg("正在优化 {0}", filePath);
                    TableExcelData data = TableExcelReader.loadFromExcel(filePath);
                    TableExcelWriter.genExcel(data, filePath);
                }
                catch (System.Exception ex)
                {
                    Log.Err(ex.Message);
                }
            }
            Log.Msg("优化完毕！");
        }

        public void genSingleFile(string filePath, string exportDir, TableExportFormat fmt)
        {
            if (!File.Exists(filePath))
            {
                Log.Err("源文件不存在！ {0}", filePath);
                return;
            }

            if (!Directory.Exists(exportDir))
            {
                Log.Err("导出目录不存在！ {0}", exportDir);
                return;
            }

            if (fmt == TableExportFormat.Unknown)
            {
                Log.Err("导出数据格式不合法！");
                return;
            }

            Log.Msg("=================================================");
            _genSingleFileImpl(filePath, exportDir, fmt);
        }

        private bool _genSingleFileImpl(string filePath, string exportDir, TableExportFormat fmt)
        {
            Log.Msg("----------正在导出 {0}----------", filePath);
            string exportPath = exportDir;
            try
            {
                TableExcelData data = TableExcelReader.loadFromExcel(filePath);
                string errmsg;
                if (!data.checkUnique(out errmsg))
                    throw new System.Exception(errmsg);
                if (!data.checkDataType(out errmsg))
                    throw new System.Exception(errmsg);
                switch (fmt)
                {
                    case TableExportFormat.Dat:
                        {
                            exportPath = Path.Combine(exportDir, string.Format("{0}.exdat", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportDat.exportExcelFile(data, exportPath);
                        }
                        break;
                    case TableExportFormat.Json:
                        {
                            exportPath = Path.Combine(exportDir, string.Format("{0}.json", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportJson.exportExcelFile(data, exportPath);
                            string fileName = Path.GetFileNameWithoutExtension(filePath);
                            fileName = fileName.Substring(0, 1).ToUpper() + fileName.Substring(1);
                            exportPath = Path.Combine(exportDir, string.Format($"{ fileName }.cs"));
                            TableExcelExportCs.ExportExcelFile(data, exportPath);
                        }
                        break;
                    case TableExportFormat.Xml:
                        {
                            exportPath = Path.Combine(exportDir, string.Format("{0}.xml", Path.GetFileNameWithoutExtension(filePath)));
                            TableExcelExportXml.exportExcelFile(data, exportPath);
                        }
                        break;
                    case TableExportFormat.Lua:
                        {
                            exportPath = Path.Combine(exportDir, string.Format("{0}.lua", Path.GetFileNameWithoutExtension(filePath)));
                            if (AppData.Config.FitUnity3D == true)
                            {
                                exportPath += ".txt";
                            }
                            TableExcelExportLua.exportExcelFile(data, exportPath);
                        }
                        break;
                }
                Log.Msg("----------导出成功 {0} ----------", exportPath);
                return true;
            }
            catch (System.Exception ex)
            {
                Log.Err(ex.Message);
                Log.Err("----------导出失败 {0} ----------", exportPath);
                return false;
            }
        }

        private void _flipHelpPanel()
        {
            var hp = new HelperPanel();
            var pw = new PopupWindow(hp);
            pw.Owner = Window.GetWindow(this);
            pw.Title = "使用说明";
            pw.ShowDialog();
        }

        private void _flipMoreSettingPanel()
        {
            var fsm = new FrameSettingMore();
            var pw = new PopupWindow(fsm);
            pw.Owner = Window.GetWindow(this);
            pw.Title = "更多设置";
            pw.ShowDialog();
        }
    }
}
