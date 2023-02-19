using DAVISystemExporter.Utils;
using Microsoft.Win32;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Windows.Threading;

namespace DAVISystemExporter.Domain
{
    public class SerialPortItem
    {
        public string Name { get; set; } = "";
        public string Port { get; set; } = "";

        public override bool Equals(object? obj)
        {
            return obj is SerialPortItem item &&
                   Name == item.Name &&
                   Port == item.Port;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Name, Port);
        }

        public override string? ToString()
        {
            return $"{Name} ({Port})";
        }
    }

    public class LogMessage
    {
        public string Message { get; set; } = "";
        public DateTime Timestamp { get; set; } = DateTime.Now;

        public override bool Equals(object? obj)
        {
            return obj is LogMessage message &&
                   Message == message.Message &&
                   Timestamp == message.Timestamp;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Message, Timestamp);
        }

        public override string? ToString()
        {
            return $"{Timestamp.ToString("hh:mm:ss.fff tt")} : {Message}";
        }
    }

    public class ReceivedEventArgs
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string[] Values { get; set; } = { };
    }
    
    public class MainWindowViewModel : ViewModelBase, IDisposable
    {
        private ThrottleDispatcher throttleDispatcher = new ThrottleDispatcher(100);
        private Dispatcher dispatcher;

        private List<LogMessage> _backLogMessages = new List<LogMessage>();
        private ObservableCollection<LogMessage> _logMessages = new ObservableCollection<LogMessage>();
        private ObservableCollection<SerialPortItem> _serialPorts = new ObservableCollection<SerialPortItem>();
        private SerialPortItem? _selectedItem = null;
        private bool _connected = false;
        private bool _create_worker = false;
        private bool _data_parse_start = false;
        private SerialPort? _connectPort = null;
        private string _lastErrorMessage = "";
        private List<byte> _buffer = new List<byte>();

        private string _baud_rate = "115200";
        private string _data_count = "100";

        public ReadOnlyObservableCollection<SerialPortItem> SerialPorts { get => new ReadOnlyObservableCollection<SerialPortItem>(_serialPorts); }
        public ReadOnlyObservableCollection<LogMessage> LogMessages { get => new ReadOnlyObservableCollection<LogMessage>(_logMessages); }

        public delegate void ReceivedEventHandler(object sender, ReceivedEventArgs e);
        public event ReceivedEventHandler? ReceivedEvent;

        public SerialPortItem? SelectedItem
        {
            get => _selectedItem;
            set => SetProperty(ref _selectedItem, value);
        }

        public bool Connected
        {
            get => _connected;
            set => SetProperty(ref _connected, value);
        }

        public bool CreateWorker
        {
            get => _create_worker;
            set => SetProperty(ref _create_worker, value);
        }

        public bool DataParseStart
        {
            get => _data_parse_start;
            set => SetProperty(ref _data_parse_start, value);
        }

        public SerialPort? ConnectPort
        {
            get => _connectPort;
            set => SetProperty(ref _connectPort, value);
        }

        public string LastErrorMessage
        {
            get => _lastErrorMessage;
            set => SetProperty(ref _lastErrorMessage, value);
        }

        public string BaudRate
        {
            get => _baud_rate;
            set => SetProperty(ref _baud_rate, value);
        }

        public int BaudRateNumber
        {
            get
            {
                try
                {
                    if (int.TryParse(_baud_rate, out int n))
                    {
                        return n;
                    }
                }
                catch { }

                return 115200;
            }
        }

        public string DataCount
        {
            get => _data_count;
            set => SetProperty(ref _data_count, value);
        }

        public int DataCountNumber
        {
            get {
                try
                {
                    if (int.TryParse(_data_count, out int n))
                    {
                        return n;
                    }
                }
                catch { }

                return 100;
            }
        }

        public MainWindowViewModel(Dispatcher dispatcher)
        {
            this.dispatcher = dispatcher;
        }

        public void RefreshSerialPort()
        {
            var backList = _serialPorts.ToHashSet();
            var newList = new HashSet<SerialPortItem>();

            using (ManagementClass i_Entity = new ManagementClass("Win32_PnPEntity"))
            {
                foreach (ManagementObject i_Inst in i_Entity.GetInstances())
                {
                    object o_Guid = i_Inst.GetPropertyValue("ClassGuid");
                    if (o_Guid?.ToString()?.ToUpper() != "{4D36E978-E325-11CE-BFC1-08002BE10318}")
                        continue; // Skip all devices except device class "PORTS"

                    string s_Caption = i_Inst.GetPropertyValue("Caption")?.ToString() ?? "";
                    string s_Manufact = i_Inst.GetPropertyValue("Manufacturer")?.ToString() ?? "";
                    string s_DeviceID = i_Inst.GetPropertyValue("PnpDeviceID")?.ToString() ?? "";
                    string s_RegPath = "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Enum\\" + s_DeviceID + "\\Device Parameters";
                    string s_PortName = Registry.GetValue(s_RegPath, "PortName", "")?.ToString() ?? "";

                    int s32_Pos = s_Caption?.IndexOf("(COM") ?? -1;
                    if (s_Caption != null && s32_Pos > 0) // remove COM port from description
                        s_Caption = s_Caption.Substring(0, s32_Pos);

                    newList.Add(new SerialPortItem()
                    {
                        Name = s_Caption ?? "",
                        Port = s_PortName
                    });
                }
            }

            if (!backList.SetEquals(newList))
            {
                var selectItem = SelectedItem;
                _serialPorts.Clear();
                if (newList.Count > 0)
                {
                    foreach (var item in newList)
                    {
                        _serialPorts.Add(item);
                    }

                    if (selectItem != null && _serialPorts.IndexOf(selectItem) >= 0)
                    {
                        SelectedItem = selectItem;
                    }
                    else
                    {
                        SelectedItem = newList.Last();
                    }
                }
                else
                {
                    SelectedItem = null;
                }
            }
        }

        public bool Connect()
        {
            var selectedItem = SelectedItem;
            if (selectedItem == null) return false;
            if (Connected) return true;

            var port = new SerialPort(selectedItem.Port);
            port.BaudRate = BaudRateNumber;
            port.DataBits = 8;
            port.Parity = Parity.None;
            port.StopBits = StopBits.One;
            port.NewLine = "\r\n";
            port.Encoding = Encoding.UTF8;

            try
            {
                AddLog($"장치 연결중 : {selectedItem.Port}");
                port.Open();

                ConnectPort = port;
                Connected = port.IsOpen;

                port.DataReceived += Port_DataReceived;
                port.ErrorReceived += Port_ErrorReceived;

                AddLog($"장치 연결됨 : {selectedItem.Port}");
                return true;
            }
            catch (Exception ex)
            {
                LastErrorMessage = ex.ToString();
                AddLog($"장치 연결 오류 : {selectedItem.Port}, ${LastErrorMessage}");
                Close(true);
            }

            return false;
        }

        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var port = sender as SerialPort;
            if (port != null && ConnectPort == port && Connected)
            {
                var message = "";
                switch (e.EventType)
                {
                    case SerialData.Chars:

                        byte[] newLine = port.Encoding.GetBytes("\r\n");
                        int lastCheckdPostion = _buffer.Count < newLine.Length ? 0 : _buffer.Count - newLine.Length;

                        byte[] data = ArrayPool<byte>.Shared.Rent(port.BytesToRead);
                        try
                        {
                            int readBytes = port.Read(data, 0, data.Length);
                            _buffer.AddRange(data);
                            message = $"{port.Encoding.GetString(data)}";
                        }
                        finally
                        {
                            ArrayPool<byte>.Shared.Return(data);
                        }

                        if (_buffer.Count > lastCheckdPostion)
                        {
                            var events = new List<ReceivedEventArgs>();
                            int findIndex = findSequence(_buffer, lastCheckdPostion, newLine);
                            while (findIndex >= 0)
                            {
                                int endIndex = findIndex;

                                var list = _buffer.GetRange(0, endIndex);
                                _buffer.RemoveRange(0, endIndex + newLine.Length);

                                var m = port.Encoding.GetString(list.ToArray());
                                AddLog($"데이터 수집중 : {port.PortName}, {m}");

                                if (_data_parse_start)
                                {
                                    string lastSufix = "@";
                                    int startIndex = m.IndexOf("$");
                                    endIndex = m.LastIndexOf(lastSufix);
                                    if (startIndex == 0 && endIndex == m.Length - lastSufix.Length)
                                    {
                                        events.Add(new ReceivedEventArgs
                                        {
                                            Values = m.Substring(startIndex + 1, m.Length - lastSufix.Length - (startIndex + 1)).Split(",")
                                        });
                                    }
                                }

                                findIndex = findSequence(_buffer, 0, newLine);
                            }

                            if (events.Count > 0)
                            {
                                dispatcher.BeginInvoke(new Action(() =>
                                {
                                    events.ForEach(e =>
                                    {
                                        ReceivedEvent?.Invoke(this, e);
                                    });
                                }));
                            }
                        }

                        return;
                    case SerialData.Eof:
                        message = "[Eof]";
                        break;
                    default:
                        message = "[Unknouwn]";
                        break;
                }

                AddLog($"데이터 수집중 : {port.PortName}, {message}");
            }
        }

        private void Port_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            var port = sender as SerialPort;
            if (ConnectPort == port && Connected)
            {
                switch (e.EventType)
                {
                    case SerialError.Frame:
                        LastErrorMessage = "Framing error ";
                        break;
                    case SerialError.Overrun:
                        LastErrorMessage = "character-buffer overrun ";
                        break;
                    case SerialError.RXOver:
                        LastErrorMessage = "Input buffer overflow";
                        break;
                    case SerialError.RXParity:
                        LastErrorMessage = "parity error pada hardware";
                        break;
                    case SerialError.TXFull:
                        LastErrorMessage = "transmit data, namun output buffer sedang penuh";
                        break;
                    default:
                        LastErrorMessage = "Unknouwn";
                        break;
                }

                AddLog($"오류 발생 : {port?.PortName}, {LastErrorMessage}");

                dispatcher.BeginInvoke(new Action(() =>
                {
                    Close(true);
                }));
            }
        }

        public void Close(bool isError = false)
        {
            DataParseStart = false;

            var connectPort = ConnectPort;
            ConnectPort = null;
            if (connectPort != null)
            {
                connectPort.DataReceived -= Port_DataReceived;
                connectPort.ErrorReceived -= Port_ErrorReceived;

                connectPort.Close();

                if (!isError)
                {
                    AddLog($"장치 연결 해제됨 : {connectPort.PortName}");
                }
            }

            Connected = false;
        }

        public void Dispose()
        {
            Close();
        }

        private void AddLog(string message)
        {
            lock (_backLogMessages)
            {
                _backLogMessages.Add(new LogMessage
                {
                    Message = message.TrimEnd()
                });

                if (_backLogMessages.Count > 10000)
                {
                    foreach (var item in _backLogMessages.Take(_backLogMessages.Count - 10000))
                    {
                        _backLogMessages.Remove(item);
                    }
                }
            }

            throttleDispatcher.Throttle(() =>
            {
                var backLogMessages = new List<LogMessage>(_backLogMessages.Count);
                lock (_backLogMessages)
                {
                    backLogMessages.AddRange(_backLogMessages);
                }

                dispatcher.BeginInvoke(new Action(() =>
                {
                    int i = 0;
                    if (backLogMessages.Count == _logMessages.Count)
                    {
                        foreach (var item in backLogMessages)
                        {
                            _logMessages[i++] = item;
                        }
                    }
                    else
                    {
                        foreach (var item in backLogMessages)
                        {
                            if (_logMessages.Count > i + 1)
                            {
                                _logMessages[i++] = item;
                            }
                            else
                            {
                                _logMessages.Add(item);
                            }
                        }
                    }
                }));
            });
        }

        /// <summary>Looks for the next occurrence of a sequence in a byte array</summary>
        /// <param name="array">Array that will be scanned</param>
        /// <param name="start">Index in the array at which scanning will begin</param>
        /// <param name="sequence">Sequence the array will be scanned for</param>
        /// <returns>
        ///   The index of the next occurrence of the sequence of -1 if not found
        /// </returns>
        private static int findSequence(List<byte> array, int start, byte[] sequence)
        {
            int end = array.Count - sequence.Length; // past here no match is possible
            byte firstByte = sequence[0]; // cached to tell compiler there's no aliasing

            while (start <= end)
            {
                // scan for first byte only. compiler-friendly.
                if (array[start] == firstByte)
                {
                    // scan for rest of sequence
                    for (int offset = 1; ; ++offset)
                    {
                        if (offset == sequence.Length)
                        { // full sequence matched?
                            return start;
                        }
                        else if (array[start + offset] != sequence[offset])
                        {
                            break;
                        }
                    }
                }
                ++start;
            }

            // end of array reached without match
            return -1;
        }
    }
}
