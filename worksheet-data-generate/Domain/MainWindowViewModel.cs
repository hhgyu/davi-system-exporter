using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Windows.Threading;

namespace worksheet_data_generate.Domain
{
    public class SerialPortItem
    {
        public string Name { get; set; }
        public string Port { get; set; }

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
        public string Message { get; set; }
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
        public string[] Values { get; set; }
    }
    
    public class MainWindowViewModel : ViewModelBase, IDisposable
    {
        private Dispatcher dispatcher;

        private ObservableCollection<LogMessage> _logMessages = new ObservableCollection<LogMessage>();
        private ObservableCollection<SerialPortItem> _serialPorts = new ObservableCollection<SerialPortItem>();
        private SerialPortItem? _selectedItem = null;
        private bool _connected = false;
        private bool _dataParseStart = false;
        private SerialPort? _connectPort = null;
        private string _lastErrorMessage = "";

        public ReadOnlyObservableCollection<SerialPortItem> SerialPorts { get => new ReadOnlyObservableCollection<SerialPortItem>(_serialPorts); }
        public ReadOnlyObservableCollection<LogMessage> LogMessages { get => new ReadOnlyObservableCollection<LogMessage>(_logMessages); }

        public delegate void ReceivedEventHandler(object sender, ReceivedEventArgs e);
        public event ReceivedEventHandler ReceivedEvent;

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

        public bool DataParseStart
        {
            get => _dataParseStart;
            set => SetProperty(ref _dataParseStart, value);
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
            port.BaudRate = 115200;
            port.DataBits = 8;
            port.Parity = Parity.None;
            port.StopBits = StopBits.One;
            port.NewLine = "\r\n";

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
            if (ConnectPort == port && Connected)
            {
                var message = "";
                switch (e.EventType)
                {
                    case SerialData.Chars:
                        message = $"{port.ReadExisting()}";

                        if(_dataParseStart)
                        {
                            foreach(var m in message.Split("\r\n"))
                            {
                                string lastSufix = "@";
                                int startIndex = m.IndexOf("$");
                                int endIndex = m.LastIndexOf(lastSufix);
                                if (startIndex == 0 && endIndex == m.Length - lastSufix.Length)
                                {
                                    dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        ReceivedEvent?.Invoke(this, new ReceivedEventArgs
                                        {
                                            Values = m.Substring(startIndex + 1, m.Length - lastSufix.Length - (startIndex + 1)).Split(",")
                                        });
                                    }));
                                }
                            }
                        }
                        break;
                    case SerialData.Eof:
                        message = "[Eof]";
                        break;
                    default:
                        message = "[Unknouwn]";
                        break;
                }

                dispatcher.BeginInvoke(new Action(() =>
                {
                    AddLog($"데이터 수집중 : {port.PortName}, {message}");
                }));
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

                dispatcher.BeginInvoke(new Action(() =>
                {
                    AddLog($"오류 발생 : {port.PortName}, {LastErrorMessage}");

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
            _logMessages.Add(new LogMessage
            {
                Message = message.TrimEnd()
            });

            if (_logMessages.Count > 3000)
            {
                foreach(var item in _logMessages.Take(_logMessages.Count - 3000)) {
                    _logMessages.Remove(item);
                }
            }
        }
    }
}
