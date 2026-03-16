using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using ImageMagick;

namespace ScreenToGifTool
{
    public partial class MainWindow : Window
    {
        private Rect _selectedRect;
        private SelectionWindow _selectionWin;
        private DispatcherTimer _recordTimer;
        private List<MagickImage> _frames = new List<MagickImage>();
        private string _savePath = @"C:\RecordedGifs";

        public MainWindow()
        {
            InitializeComponent();
            // 程式啟動時自動建立資料夾
            if (!Directory.Exists(_savePath)) Directory.CreateDirectory(_savePath);
        }

        private void BtnSelectArea_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            SelectionWindow selectWin = new SelectionWindow();

            // 再次強制確保它蓋住所有螢幕範圍
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
            _frames.Clear();
            int fps = cbQuality.SelectedIndex == 0 ? 30 : 60;

            _recordTimer = new DispatcherTimer(DispatcherPriority.Render);
            _recordTimer.Interval = TimeSpan.FromMilliseconds(1000.0 / fps);
            _recordTimer.Tick += CaptureScreen;

            btnStart.IsEnabled = false;
            btnStop.IsEnabled = true;
            btnSelectArea.IsEnabled = false;

            // --- 修正版：建立提醒框 ---
            _selectionWin = new SelectionWindow();
            _selectionWin.IsHitTestVisible = false;

            // 關鍵修正：讓提醒框比錄製區域寬 4 像素 (左右各 2)
            // 這樣綠線就會在外圍，不會被 CopyFromScreen 抓到
            _selectionWin.Left = _selectedRect.X - 2;
            _selectionWin.Top = _selectedRect.Y - 2;
            _selectionWin.Width = _selectedRect.Width + 4;
            _selectionWin.Height = _selectedRect.Height + 4;

            _selectionWin.SetRecordingStyle(true);

            _selectionWin.SelectionBorder.Visibility = Visibility.Visible;
            _selectionWin.SelectionBorder.Margin = new Thickness(0);
            _selectionWin.SelectionBorder.Width = _selectedRect.Width + 4;
            _selectionWin.SelectionBorder.Height = _selectedRect.Height + 4;

            _selectionWin.Show();
            _recordTimer.Start();
        }

        private void CaptureScreen(object sender, EventArgs e)
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

                    using (MemoryStream ms = new MemoryStream())
                    {
                        bmp.Save(ms, ImageFormat.Bmp);
                        ms.Position = 0;

                        var image = new MagickImage(ms);
                        // 這裡最關鍵：雖然 MagickImage 不是 WPF 物件，
                        // 但為了保險，我們確保它在加入清單前已經完全脫離 MemoryStream
                        _frames.Add(image);
                    }
                }
            }
            catch { /* 忽略跳幀 */ }
        }

        private async void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            _recordTimer?.Stop();
            if (_selectionWin != null) _selectionWin.Hide();

            pbProcessing.Visibility = Visibility.Visible;
            btnStop.IsEnabled = false;

            // 先在 UI 執行緒拿好設定值，不要在 Task.Run 裡面拿
            int qualityIndex = cbQuality.SelectedIndex;

            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    string fileName = $"Gif_{DateTime.Now:yyyyMMdd_HHmmss}.gif";
                    string fullPath = Path.Combine(_savePath, fileName);

                    using (MagickImageCollection collection = new MagickImageCollection())
                    {
                        uint delay = qualityIndex == 0 ? 3u : 1u;

                        foreach (var frame in _frames)
                        {
                            frame.AnimationDelay = delay;
                            if (qualityIndex == 0)
                            {
                                frame.Resize(frame.Width / 2u, 0u);
                            }
                            collection.Add(frame);
                        }

                        collection.Optimize();
                        collection.Write(fullPath);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        pbProcessing.Visibility = Visibility.Collapsed;
                        MessageBox.Show($"錄製完成！\n路徑：{fullPath}", "成功");
                        btnStart.IsEnabled = true;
                        btnSelectArea.IsEnabled = true;
                    });
                }
                catch (Exception ex)
                {
                    // 修正：使用 Dispatcher 彈出錯誤視窗
                    Dispatcher.Invoke(() => MessageBox.Show($"存檔失敗：{ex.Message}"));
                }
                finally
                {
                    foreach (var f in _frames) f.Dispose();
                    _frames.Clear();
                }
            });
        }
    }
}