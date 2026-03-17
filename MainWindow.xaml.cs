using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using ImageMagick;
using System.Threading.Tasks;

namespace ScreenToGifTool
{
    public partial class MainWindow : Window
    {
        private Rect _selectedRect;
        private SelectionWindow? _selectionWin;
        private DispatcherTimer? _recordTimer;

        private List<string> _frameFiles = new List<string>();

        // 【修正 1】定義存檔與暫存路徑
        private static readonly string _savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ScreenToGif_Output");
        private static readonly string _tempPath = Path.Combine(Path.GetTempPath(), "ScreenToGif_Temp");

        public MainWindow()
        {
            InitializeComponent();

            // 【修正 2】啟動時檢查並建立存檔目錄
            if (!Directory.Exists(_savePath)) Directory.CreateDirectory(_savePath);

            ResetTempDirectory();
        }

        private void ResetTempDirectory()
        {
            try
            {
                if (Directory.Exists(_tempPath)) Directory.Delete(_tempPath, true);
                Directory.CreateDirectory(_tempPath);
                _frameFiles.Clear();
            }
            catch { /* 防止檔案被佔用時崩潰 */ }
        }

        private void BtnSelectArea_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            SelectionWindow selectWin = new SelectionWindow();

            // 跨螢幕座標支援
            selectWin.Left = SystemParameters.VirtualScreenLeft;
            selectWin.Top = SystemParameters.VirtualScreenTop;
            selectWin.Width = SystemParameters.VirtualScreenWidth;
            selectWin.Height = SystemParameters.VirtualScreenHeight;

            bool? result = selectWin.ShowDialog();
            if (result == true)
            {
                _selectedRect = selectWin.SelectedRect;
                btnStart.IsEnabled = true;
            }
            this.Show();
        }

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            ResetTempDirectory();

            int fps = cbQuality.SelectedIndex == 0 ? 30 : 60;
            _recordTimer = new DispatcherTimer(DispatcherPriority.Render);
            _recordTimer.Interval = TimeSpan.FromMilliseconds(1000.0 / fps);
            _recordTimer.Tick += CaptureScreen;

            btnStart.IsEnabled = false;
            btnStop.IsEnabled = true;
            btnSelectArea.IsEnabled = false;

            // 建立錄影提醒框
            _selectionWin = new SelectionWindow();
            _selectionWin.IsHitTestVisible = false;

            // 綠框往外擴張 2 像素避免被錄進去
            _selectionWin.Left = _selectedRect.X - 2;
            _selectionWin.Top = _selectedRect.Y - 2;
            _selectionWin.Width = _selectedRect.Width + 4;
            _selectionWin.Height = _selectedRect.Height + 4;

            _selectionWin.SetRecordingStyle(true);
            _selectionWin.SelectionBorder.Visibility = Visibility.Visible;
            _selectionWin.SelectionBorder.Width = _selectedRect.Width + 4;
            _selectionWin.SelectionBorder.Height = _selectedRect.Height + 4;
            _selectionWin.Show();

            _recordTimer.Start();
            lblStatus.Content = "正在錄製中...";
        }

        private void CaptureScreen(object? sender, EventArgs e)
        {
            try
            {
                int x = (int)_selectedRect.X;
                int y = (int)_selectedRect.Y;
                int w = (int)_selectedRect.Width;
                int h = (int)_selectedRect.Height;

                using (Bitmap bmp = new Bitmap(w, h))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.CopyFromScreen(x, y, 0, 0, bmp.Size);
                    }

                    // 快速存入硬碟暫存
                    string fileName = Path.Combine(_tempPath, $"f_{_frameFiles.Count}.jpg");
                    bmp.Save(fileName, ImageFormat.Jpeg);
                    _frameFiles.Add(fileName);
                }
            }
            catch { }
        }

        private async void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            // 1. 在 UI 執行緒先處理 UI 狀態
            _recordTimer?.Stop();
            if (_selectionWin != null) _selectionWin.Hide();

            pbProcessing.Visibility = Visibility.Visible;
            btnStop.IsEnabled = false;
            lblStatus.Content = "正在合成 GIF (這可能需要一點時間)...";

            // 【關鍵點】在進入 Task 之前，先取出 ComboBox 的值
            // 這樣背景執行緒就不會去「碰」UI 物件，也就不會報錯了
            int qualityIndex = cbQuality.SelectedIndex;

            // 2. 開始背景合成
            await Task.Run(() =>
            {
                try
                {
                    string fileName = $"Gif_{DateTime.Now:yyyyMMdd_HHmmss}.gif";
                    string finalPath = Path.Combine(_savePath, fileName);

                    using (MagickImageCollection collection = new MagickImageCollection())
                    {
                        // 修改延遲時間：數字越大越慢
                        // 10 代表 0.1秒一張圖，這對教學 GIF 來說很舒服
                        uint delay = (qualityIndex == 0) ? 10u : 6u;

                        foreach (string file in _frameFiles)
                        {
                            var img = new MagickImage(file);
                            img.AnimationDelay = delay; // 套用新的延遲

                            if (qualityIndex == 0)
                            {
                                img.Resize(img.Width / 2u, 0u);
                            }
                            collection.Add(img);
                        }

                        collection.OptimizePlus();
                        // 這裡可以加入重複播放設定 (0 代表無限循環)
                        collection[0].AnimationIterations = 0;

                        collection.Quantize(new QuantizeSettings { Colors = 256 });
                        collection.Write(finalPath);
                    }

                    // 3. 合成完畢，回到 UI 執行緒顯示結果
                    Dispatcher.Invoke(() =>
                    {
                        pbProcessing.Visibility = Visibility.Collapsed;
                        lblStatus.Content = "錄製完成！";

                        // 自動開啟資料夾
                        try
                        {
                            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{finalPath}\"");
                        }
                        catch
                        {
                            System.Diagnostics.Process.Start("explorer.exe", _savePath);
                        }

                        MessageBox.Show($"錄製成功！\n檔案位置：{finalPath}");
                        btnStart.IsEnabled = true;
                        btnSelectArea.IsEnabled = true;
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => {
                        lblStatus.Content = "合成失敗";
                        MessageBox.Show($"合成失敗：{ex.Message}");
                        btnStart.IsEnabled = true;
                        btnSelectArea.IsEnabled = true;
                    });
                }
            });
        }


    }
}