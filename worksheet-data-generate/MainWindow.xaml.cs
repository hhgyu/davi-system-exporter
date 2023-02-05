using MaterialDesignExtensions.Controls;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Specialized;
using System.IO;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using worksheet_data_generate.Domain;
using worksheet_data_generate.Extensions;

namespace worksheet_data_generate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MaterialWindow
    {
        private readonly MainWindowViewModel _viewModel;

        public const string DialogHostName = "RootDialog";

        IExcelWorker? _worker = null;

        bool scrollLock = false;

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
                    // scroll the new item into view   
                    LogListBox.ScrollIntoView(e.NewItems[0]);
                }
            }
        }

        private void NewExcel_Click(object sender, RoutedEventArgs e)
        {
            _worker?.Dispose();

            _worker = new ExcelWorker();
            _worker.Create();
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
                _worker?.Dispose();

                _worker = new ExcelWorker();
                string errorMessage = "";
                if (!_worker.Open(result.File, out errorMessage))
                {
                    await DialogHost.Show(errorMessage, dialogIdentifier: DialogHostName);
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
                var worker = new ExcelWorker();
                var result = worker.Check();
                worker.Dispose();
                
                /*
                IExcelWorker _worker2 = new ExcelWorker();
                _worker2.Create();
                _worker2.Start();
                MessageBox.Show("1");

                await Task.Delay(1000);

                for (int k= 0; k < 1000; k++)
                {
                    for (int i = 0; i < 1000; i++)
                    {
                        _worker2.Enqueue(DateTime.Now, new string[] { ((k * 1000) + i).ToString(), ((k * 1000) + i + 1).ToString(), ((k * 1000) + i + 2).ToString(), ((k * 1000) + i + 3).ToString() });
                    }

                    await Task.Delay(1000);
                }

                MessageBox.Show("asdfasf");

                _worker2.Dispose();
                */

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

            _viewModel.Close();

            _worker?.Dispose();
            _worker = null;
        }

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
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
    }
}
