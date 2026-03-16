using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ScreenToGifTool
{
    public partial class SelectionWindow : Window
    {
        private Point _startPoint;
        private bool _isSelecting = false;
        private bool _isFixed = false;

        // 用來給 MainWindow 讀取的錄影區域資訊
        public Rect SelectedRect { get; private set; }

        public SelectionWindow()
        {
            InitializeComponent();

            // 關鍵：抓取所有螢幕組合起來的「虛擬螢幕」範圍
            this.Left = SystemParameters.VirtualScreenLeft;
            this.Top = SystemParameters.VirtualScreenTop;
            this.Width = SystemParameters.VirtualScreenWidth;
            this.Height = SystemParameters.VirtualScreenHeight;

            // 確保視窗在最前面
            this.Topmost = true;
            this.WindowStartupLocation = WindowStartupLocation.Manual;
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                if (!_isSelecting && !_isFixed)
                {
                    // 第一下：開始框選
                    _isSelecting = true;
                    _startPoint = e.GetPosition(this);
                }
                else if (_isSelecting)
                {
                    // 第二下：固定框選
                    _isSelecting = false;
                    _isFixed = true;

                    // 彈出確認視窗
                    var result = MessageBox.Show("確認選取此區域嗎？", "確認", MessageBoxButton.OKCancel);

                    if (result == MessageBoxResult.OK)
                    {
                        // 用戶點「確定」：關閉視窗並傳回 true
                        this.DialogResult = true;
                    }
                    else
                    {
                        // 用戶點「取消」：關閉視窗並傳回 false
                        // 這會讓 MainWindow 裡的 ShowDialog 結束，並執行之後的 this.Show()
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
                double width = Math.Abs(_startPoint.X - currentPoint.X);
                double height = Math.Abs(_startPoint.Y - currentPoint.Y);

                // 顯示框框
                SelectionBorder.Visibility = Visibility.Visible;
                SelectionBorder.Width = width;
                SelectionBorder.Height = height;

                // 使用 Canvas 設定座標
                Canvas.SetLeft(SelectionBorder, x);
                Canvas.SetTop(SelectionBorder, y);

                SelectedRect = new Rect(x, y, width, height);
            }
        }

        // 當錄製開始時，MainWindow 會呼叫這個方法來變色
        public void SetRecordingStyle(bool isRecording)
        {
            SelectionBorder.BorderBrush = isRecording ? Brushes.Green : Brushes.Red;
        }
    }
}