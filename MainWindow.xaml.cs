using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using ImageMagick;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ScreenToGifTool
{
    public partial class MainWindow : Window
    {
        private Rect _selectedRect;
        private SelectionWindow? _selectionWin;
        private DispatcherTimer? _recordTimer;
        private List<string> _frameFiles = new List<string>();

        private static readonly string _savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ScreenToGif_Output");
        private static readonly string _tempPath = Path.Combine(Path.GetTempPath(), "ScreenToGif_Temp");

        public MainWindow()
        {
            InitializeComponent();
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
            catch { }
        }

        private void BtnSelectArea_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            SelectionWindow selectWin = new SelectionWindow();
            selectWin.Left = SystemParameters.VirtualScreenLeft;
            selectWin.Top = SystemParameters.VirtualScreenTop;
            selectWin.Width = SystemParameters.VirtualScreenWidth;
            selectWin.Height = SystemParameters.VirtualScreenHeight;

            if (selectWin.ShowDialog() == true)
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

            _selectionWin = new SelectionWindow();
            _selectionWin.Show();
            _selectionWin.EnableClickThrough();

            _selectionWin.Left = _selectedRect.X - 2;
            _selectionWin.Top = _selectedRect.Y - 2;
            _selectionWin.Width = _selectedRect.Width + 4;
            _selectionWin.Height = _selectedRect.Height + 4;

            _selectionWin.SetRecordingStyle(true);
            _selectionWin.SelectionBorder.Visibility = Visibility.Visible;
            _selectionWin.SelectionBorder.Width = _selectedRect.Width + 4;
            _selectionWin.SelectionBorder.Height = _selectedRect.Height + 4;

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
                    string fileName = Path.Combine(_tempPath, $"f_{_frameFiles.Count}.jpg");
                    bmp.Save(fileName, ImageFormat.Jpeg);
                    _frameFiles.Add(fileName);
                }
            }
            catch { }
        }


        private async void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            // 1. 停止錄影並徹底關閉選取視窗，避免進程殘留
            _recordTimer?.Stop();
            if (_selectionWin != null)
            {
                _selectionWin.Close();
                _selectionWin = null;
            }

            // 2. 更新 UI 狀態
            pbProcessing.Visibility = Visibility.Visible;
            btnStop.IsEnabled = false;
            lblStatus.Content = "正在合成 GIF (請稍候)...";

            // 3. 提取背景執行緒所需的「純資料」變數 (不直接存取 UI 元件)
            int currentQualityIndex = cbQuality.SelectedIndex;
            List<string> framesToProcess = new List<string>(_frameFiles);
            string saveDir = _savePath;

            // 4. 開始非同步合成，避免 UI 凍結
            await Task.Run(() =>
            {
                try
                {
                    string fileName = $"Gif_{DateTime.Now:yyyyMMdd_HHmmss}.gif";
                    string finalPath = Path.Combine(saveDir, fileName);

                    using (MagickImageCollection collection = new MagickImageCollection())
                    {
                        // 【速度修正】
                        // 30 FPS 錄製應對應 3u 或 4u (這裡用 4u 體感最自然)
                        // 60 FPS 錄製應對應 2u (GIF 物理極限順暢度)
                        uint delay = (currentQualityIndex == 0) ? 6u : 4u;

                        foreach (string file in framesToProcess)
                        {
                            if (!File.Exists(file)) continue;
                            var img = new MagickImage(file);
                            img.AnimationDelay = delay;

                            // 低畫質模式下進行解析度減半壓縮
                            if (currentQualityIndex == 0)
                            {
                                img.Resize(img.Width / 2u, 0u);
                            }
                            collection.Add(img);
                        }

                        // GIF 優化處理
                        collection.OptimizePlus();
                        if (collection.Count > 0)
                        {
                            collection[0].AnimationIterations = 0; // 無限循環
                        }

                        collection.Quantize(new QuantizeSettings { Colors = 256 });
                        collection.Write(finalPath);
                    }

                    // 5. 回到 UI 執行緒處理完成狀態
                    Dispatcher.Invoke(() =>
                    {
                        pbProcessing.Visibility = Visibility.Collapsed;
                        lblStatus.Content = "錄製完成！";

                        // 【安全性優化】使用更溫和的方式開啟資料夾，減少防毒軟體誤判
                        try
                        {
                            // 嘗試開啟並選取檔案
                            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{finalPath}\"") { UseShellExecute = true });
                        }
                        catch
                        {
                            // 若失敗則僅開啟目標資料夾
                            Process.Start(new ProcessStartInfo(saveDir) { UseShellExecute = true });
                        }

                        MessageBox.Show($"錄製成功！\n檔案位置：{finalPath}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        btnStart.IsEnabled = true;
                        btnSelectArea.IsEnabled = true;
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        pbProcessing.Visibility = Visibility.Collapsed;
                        lblStatus.Content = "合成失敗";
                        MessageBox.Show($"合成錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                        btnStart.IsEnabled = true;
                        btnSelectArea.IsEnabled = true;
                    });
                }
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            // 徹底清理計時器與事件
            if (_recordTimer != null)
            {
                _recordTimer.Stop();
                _recordTimer.Tick -= CaptureScreen;
            }

            // 關閉所有視窗
            if (_selectionWin != null) _selectionWin.Close();

            // 清理暫存圖片檔，不留垃圾在使用者電腦
            try
            {
                if (Directory.Exists(_tempPath))
                {
                    Directory.Delete(_tempPath, true);
                }
            }
            catch { }

            base.OnClosed(e);

            // 確保所有執行緒結束，工作管理員不留殘影
            Application.Current.Shutdown();
            Environment.Exit(0);
        }


    }
}