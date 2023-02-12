using MaterialDesignExtensions.Controls;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Specialized;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using worksheet_data_generate.Domain;
using worksheet_data_generate.Extensions;
using worksheet_data_generate.Utils;

namespace worksheet_data_generate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MaterialWindow
    {
        public const string DialogHostName = "RootDialog";

        private readonly MainWindowViewModel _viewModel;
        private ThrottleDispatcher _throttleDispatcher = new ThrottleDispatcher(50);

        private IExcelWorker? _worker = null;

        private bool scrollLock = false;
        private bool scrollEventWorking = false;

        public MainWindow()
        {
            _viewModel = new MainWindowViewModel(Dispatcher);

            InitializeComponent();

            DataContext = _viewModel;
        }

        private void MainWindow_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if(e.Action == NotifyCollectionChangedAction.Add && e.NewItems?.Count > 0)
            {
                if(!scrollLock)
                {
                    scrollEventWorking = true;
                    // scroll the new item into view
                    _throttleDispatcher.Throttle(() =>
                    {
                        LogListBox.Dispatcher.BeginInvoke(() =>
                        {
                            LogListBox.ScrollIntoView(e.NewItems[0]);

                            scrollEventWorking = false;
                        });
                    });
                }
            }
        }

        private void NewExcel_Click(object sender, RoutedEventArgs e)
        {
            var worker = _worker;
            if(worker != null)
            {
                _viewModel.CreateWorker = false;
                worker.Stoped += null;
                worker.Dispose();
            }

            _worker = new ExcelWorker(_viewModel.DataCountNumber);
            _worker.Stoped += Worker_Stoped;
            
            _worker.Create();
            _viewModel.CreateWorker = true;
        }

        private void Worker_Stoped(object? sender, EventArgs e)
        {
        }

        private async void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            // open file
            OpenFileDialogArguments dialogArgs = new OpenFileDialogArguments()
            {
                Width = 600,
                Height = 400,
                Filters = "Excel Files(*.xlsx;*.db)|*.xlsx;*.db",
                CurrentDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
            };

            OpenFileDialogResult result = await OpenFileDialog.ShowDialogAsync(DialogHostName, dialogArgs);

            if(result.Confirmed && result.File != null)
            {
                var worker = _worker;
                if (worker != null)
                {
                    _viewModel.CreateWorker = false;
                    worker.Stoped += null;
                    worker.Dispose();
                }

                _worker = new ExcelWorker(_viewModel.DataCountNumber);
                _worker.Stoped += Worker_Stoped;
                if (!_worker.Open(result.File, out string errorMessage))
                {
                    await DialogHost.Show(errorMessage, dialogIdentifier: DialogHostName);
                }
                else
                {
                    _viewModel.CreateWorker = true;
                }
            }
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _viewModel.ReceivedEvent += _viewModel_ReceivedEvent;
            _viewModel.RefreshSerialPort();
            ((INotifyCollectionChanged)_viewModel.LogMessages).CollectionChanged += MainWindow_CollectionChanged;

            var scrollViewer = LogListBox.FindVisualDescendant<ScrollViewer>();
            if(scrollViewer != null)
            {
                scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
            }

            try
            {
                var worker = new ExcelWorker(0);
                var result = worker.Check();
                worker.Dispose();

                if (!result)
                {
                    return;
                }
            }
            catch (Exception) { }

            await DialogHost.Show("엑셀이 확인 되지 않습니다", dialogIdentifier: DialogHostName);
            Close();
        }

        private void _viewModel_ReceivedEvent(object sender, ReceivedEventArgs e)
        {
            if(e.Values.Length == _worker?.ValueRows)
            {

            }
            _worker?.Enqueue(e.Timestamp, e.Values);
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            ((INotifyCollectionChanged)_viewModel.LogMessages).CollectionChanged -= MainWindow_CollectionChanged;

            var scrollViewer = LogListBox.FindVisualDescendant<ScrollViewer>();
            if (scrollViewer != null)
            {
                scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
            }

            _worker?.Stop();

            _viewModel.Close();

            _worker?.Dispose();
            _worker = null;
        }

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if(scrollEventWorking == true)
            {
                return;
            }

            if (e.VerticalChange != 0)
            {
                var scrollViewer = sender as ScrollViewer;
                if (scrollViewer != null)
                {
                    if (scrollLock)
                    {
                        if (e.ExtentHeight - (e.VerticalOffset + e.ViewportHeight) < 2)
                        {
                            scrollLock = false;
                        }
                    }
                    else
                    {
                        if (e.ExtentHeight - (e.VerticalOffset + e.ViewportHeight) >= 3)
                        {
                            scrollLock = true;
                        }
                    }
                }
            }
        }

        private void RefreshSerialPort_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.RefreshSerialPort();
        }

        private void SerialPortConnect_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.Connected)
            {
                _worker?.Stop();

                _viewModel.Close();
            }
            else
            {
                if (_viewModel.SelectedItem == null)
                {
                    if (Snackbar.MessageQueue is { } messageQueue)
                    {
                        Task.Factory.StartNew(() => messageQueue.Enqueue("포트를 선택해주세요!"));
                    }

                    return;
                }

                if (!_viewModel.Connect())
                {
                    if (Snackbar.MessageQueue is { } messageQueue)
                    {
                        Task.Factory.StartNew(() => messageQueue.Enqueue($"연결에 실패했습니다.\r\n{_viewModel.LastErrorMessage ?? "알수없는 오류가 발생했습니다."}"));
                    }

                    return;
                }
            }
        }

        private void DataParseStart_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.DataParseStart = !_viewModel.DataParseStart;
            if (_viewModel.DataParseStart)
            {
                _worker?.Start();
            }
            else
            {
                _worker?.Stop();
            }
        }

        private void ExportData_Click(object sender, RoutedEventArgs e)
        {
            _worker?.Export();
        }
    }
}
