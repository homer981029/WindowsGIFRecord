using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls; // 必須引用這個才能用 Canvas
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace ScreenToGifTool
{
    public partial class SelectionWindow : Window
    {
        // --- Win32 API 穿透常數 ---
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TRANSPARENT = 0x00000020;

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        // --- 狀態變數 ---
        private Point _startPoint;
        private bool _isSelecting = false; // 補上這個
        private bool _isFixed = false;     // 補上這個
        public Rect SelectedRect { get; set; }

        public SelectionWindow()
        {
            InitializeComponent();
        }

        // 開啟穿透
        public void EnableClickThrough()
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_TRANSPARENT);
        }

        // 關閉穿透
        public void DisableClickThrough()
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle & ~WS_EX_TRANSPARENT);
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                if (!_isSelecting && !_isFixed)
                {
                    _isSelecting = true;
                    _startPoint = e.GetPosition(this);
                }
                else if (_isSelecting)
                {
                    _isSelecting = false;
                    _isFixed = true;

                    var result = MessageBox.Show("確認選取此區域嗎？", "確認", MessageBoxButton.OKCancel);
                    if (result == MessageBoxResult.OK)
                    {
                        this.DialogResult = true;
                    }
                    else
                    {
                        // 取消後重置狀態，讓使用者可以重新選取
                        _isFixed = false;
                        SelectionBorder.Visibility = Visibility.Collapsed;
                        this.DialogResult = false;
                    }
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_isSelecting)
            {
                Point currentPoint = e.GetPosition(this);

                double x = Math.Min(_startPoint.X, currentPoint.X);
                double y = Math.Min(_startPoint.Y, currentPoint.Y);
                double width = Math.Max(1, Math.Abs(_startPoint.X - currentPoint.X));
                double height = Math.Max(1, Math.Abs(_startPoint.Y - currentPoint.Y));

                SelectionBorder.Visibility = Visibility.Visible;
                SelectionBorder.Width = width;
                SelectionBorder.Height = height;

                Canvas.SetLeft(SelectionBorder, x);
                Canvas.SetTop(SelectionBorder, y);

                SelectedRect = new Rect(x, y, width, height);
            }
        }

        public void SetRecordingStyle(bool isRecording)
        {
            SelectionBorder.BorderBrush = isRecording ? Brushes.Green : Brushes.Red;
        }
    }
}