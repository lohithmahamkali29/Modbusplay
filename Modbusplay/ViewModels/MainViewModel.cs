using System.Collections.ObjectModel;
using System.IO.Ports;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Modbusplay.Helpers;
using Modbusplay.Models;
using Modbusplay.Services;

namespace Modbusplay.ViewModels;

public class MainViewModel : ObservableObject, IDisposable
{
    private readonly ModbusTransport _transport = new();
    private readonly Dispatcher _dispatcher;
    private CancellationTokenSource? _slaveCts;
    private DispatcherTimer? _healthTimer;
    private int _failedHealthChecks;

    public MainViewModel()
    {
        _dispatcher = Application.Current.Dispatcher;

        _transport.DataSent += data => AddLog("TX ?", data);
        _transport.DataReceived += data => AddLog("RX ?", data);
        _transport.ConnectionLost += reason => _dispatcher.BeginInvoke(() =>
        {
            if (!IsConnected) return;
            AutoDisconnect($"Connection lost: {reason}");
        });

        ConnectCommand = new RelayCommand(_ => Connect(), _ => !IsConnected);
        DisconnectCommand = new RelayCommand(_ => Disconnect(), _ => IsConnected);
        RefreshPortsCommand = new RelayCommand(_ => RefreshPorts());
        ExecuteMasterCommand = new AsyncRelayCommand(_ => ExecuteMaster(), _ => IsConnected, HandleCommandError);
        StartSlaveCommand = new AsyncRelayCommand(_ => StartSlave(), _ => !IsSlaveRunning, HandleCommandError);
        StopSlaveCommand = new RelayCommand(_ => StopSlave(), _ => IsSlaveRunning);
        AddSlaveRegisterCommand = new RelayCommand(_ => AddSlaveRegister());
        RemoveSlaveRegisterCommand = new RelayCommand(_ => RemoveSlaveRegister());
        SendCustomFrameCommand = new AsyncRelayCommand(_ => SendCustomFrame(), _ => IsConnected, HandleCommandError);
        ClearLogCommand = new RelayCommand(_ => FrameLog.Clear());

        RefreshPorts();

        FunctionCodes =
        [
            new(ModbusFunctionCode.ReadCoils, "01 - Read Coils"),
            new(ModbusFunctionCode.ReadDiscreteInputs, "02 - Read Discrete Inputs"),
            new(ModbusFunctionCode.ReadHoldingRegisters, "03 - Read Holding Registers"),
            new(ModbusFunctionCode.ReadInputRegisters, "04 - Read Input Registers"),
            new(ModbusFunctionCode.WriteSingleCoil, "05 - Write Single Coil"),
            new(ModbusFunctionCode.WriteSingleRegister, "06 - Write Single Register"),
            new(ModbusFunctionCode.WriteMultipleCoils, "15 - Write Multiple Coils"),
            new(ModbusFunctionCode.WriteMultipleRegisters, "16 - Write Multiple Registers")
        ];
        SelectedFunctionCode = FunctionCodes[2];
    }

    #region Connection Properties

    public List<string> Protocols { get; } = ["RTU", "ASCII", "TCP"];

    private string _selectedProtocol = "TCP";
    public string SelectedProtocol
    {
        get => _selectedProtocol;
        set
        {
            if (SetProperty(ref _selectedProtocol, value))
            {
                OnPropertyChanged(nameof(IsSerialMode));
                OnPropertyChanged(nameof(IsTcpMode));
            }
        }
    }

    public bool IsSerialMode => SelectedProtocol is "RTU" or "ASCII";
    public bool IsTcpMode => SelectedProtocol == "TCP";

    public ObservableCollection<string> AvailableComPorts { get; } = [];

    private string _selectedComPort = "";
    public string SelectedComPort
    {
        get => _selectedComPort;
        set => SetProperty(ref _selectedComPort, value);
    }

    public List<int> BaudRates { get; } = [300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200];

    private int _selectedBaudRate = 9600;
    public int SelectedBaudRate
    {
        get => _selectedBaudRate;
        set => SetProperty(ref _selectedBaudRate, value);
    }

    public List<int> DataBitsOptions { get; } = [7, 8];

    private int _selectedDataBits = 8;
    public int SelectedDataBits
    {
        get => _selectedDataBits;
        set => SetProperty(ref _selectedDataBits, value);
    }

    public List<string> StopBitsOptions { get; } = ["1", "1.5", "2"];

    private string _selectedStopBits = "1";
    public string SelectedStopBits
    {
        get => _selectedStopBits;
        set => SetProperty(ref _selectedStopBits, value);
    }

    public List<string> ParityOptions { get; } = ["None", "Odd", "Even", "Mark", "Space"];

    private string _selectedParity = "None";
    public string SelectedParity
    {
        get => _selectedParity;
        set => SetProperty(ref _selectedParity, value);
    }

    private string _ipAddress = "127.0.0.1";
    public string IpAddress
    {
        get => _ipAddress;
        set => SetProperty(ref _ipAddress, value);
    }

    private int _tcpPort = 502;
    public int TcpPort
    {
        get => _tcpPort;
        set => SetProperty(ref _tcpPort, value);
    }

    private byte _slaveId = 1;
    public byte SlaveId
    {
        get => _slaveId;
        set => SetProperty(ref _slaveId, value);
    }

    private int _timeoutMs = 1000;
    public int TimeoutMs
    {
        get => _timeoutMs;
        set => SetProperty(ref _timeoutMs, value);
    }

    private bool _isConnected;
    public bool IsConnected
    {
        get => _isConnected;
        set
        {
            if (SetProperty(ref _isConnected, value))
                CommandManager.InvalidateRequerySuggested();
        }
    }

    #endregion

    #region Master Properties

    public List<FunctionCodeItem> FunctionCodes { get; }

    private FunctionCodeItem _selectedFunctionCode = null!;
    public FunctionCodeItem SelectedFunctionCode
    {
        get => _selectedFunctionCode;
        set
        {
            if (SetProperty(ref _selectedFunctionCode, value))
            {
                OnPropertyChanged(nameof(IsWriteFunction));
                OnPropertyChanged(nameof(IsReadFunction));
                OnPropertyChanged(nameof(IsSingleWrite));
                OnPropertyChanged(nameof(IsMultipleWrite));
                OnPropertyChanged(nameof(ShowQuantity));
                OnPropertyChanged(nameof(IsRegisterFunction));
            }
        }
    }

    public bool IsWriteFunction => SelectedFunctionCode?.IsWrite ?? false;
    public bool IsReadFunction => !IsWriteFunction;
    public bool IsSingleWrite => SelectedFunctionCode?.IsSingleWrite ?? false;
    public bool IsMultipleWrite => SelectedFunctionCode?.IsMultipleWrite ?? false;
    public bool ShowQuantity => IsReadFunction || IsMultipleWrite;
    public bool IsRegisterFunction => !(SelectedFunctionCode?.IsCoilFunction ?? false);

    private ushort _startAddress;
    public ushort StartAddress
    {
        get => _startAddress;
        set => SetProperty(ref _startAddress, value);
    }

    private ushort _quantity = 10;
    public ushort Quantity
    {
        get => _quantity;
        set => SetProperty(ref _quantity, value);
    }

    private string _writeValues = "";
    public string WriteValues
    {
        get => _writeValues;
        set => SetProperty(ref _writeValues, value);
    }

    public List<string> ByteOrderOptions { get; } = ["MSD First (Big-Endian)", "LSD First (Little-Endian)"];

    private string _selectedByteOrder = "MSD First (Big-Endian)";
    public string SelectedByteOrder
    {
        get => _selectedByteOrder;
        set => SetProperty(ref _selectedByteOrder, value);
    }

    private ByteOrder CurrentByteOrder => _selectedByteOrder.StartsWith("MSD") ? ByteOrder.MSDFirst : ByteOrder.LSDFirst;

    public ObservableCollection<RegisterResult> RegisterResults { get; } = [];
    public ObservableCollection<CoilResult> CoilResults { get; } = [];

    private bool _showRegisterResults;
    public bool ShowRegisterResults
    {
        get => _showRegisterResults;
        set => SetProperty(ref _showRegisterResults, value);
    }

    private bool _showCoilResults;
    public bool ShowCoilResults
    {
        get => _showCoilResults;
        set => SetProperty(ref _showCoilResults, value);
    }

    private string _responseInfo = "";
    public string ResponseInfo
    {
        get => _responseInfo;
        set => SetProperty(ref _responseInfo, value);
    }

    #endregion

    #region Slave Properties

    public ObservableCollection<SlaveDataEntry> SlaveRegisters { get; } = [];

    private bool _isSlaveRunning;
    public bool IsSlaveRunning
    {
        get => _isSlaveRunning;
        set
        {
            if (SetProperty(ref _isSlaveRunning, value))
                CommandManager.InvalidateRequerySuggested();
        }
    }

    private string _slaveStatus = "Stopped";
    public string SlaveStatus
    {
        get => _slaveStatus;
        set => SetProperty(ref _slaveStatus, value);
    }

    #endregion

    #region Normal Mode Properties

    private string _customFrameHex = "";
    public string CustomFrameHex
    {
        get => _customFrameHex;
        set => SetProperty(ref _customFrameHex, value);
    }

    private bool _autoAddChecksum = true;
    public bool AutoAddChecksum
    {
        get => _autoAddChecksum;
        set => SetProperty(ref _autoAddChecksum, value);
    }

    private bool _waitForResponse = true;
    public bool WaitForResponse
    {
        get => _waitForResponse;
        set => SetProperty(ref _waitForResponse, value);
    }

    private string _customResponseHex = "";
    public string CustomResponseHex
    {
        get => _customResponseHex;
        set => SetProperty(ref _customResponseHex, value);
    }

    #endregion

    #region Frame Log

    public ObservableCollection<FrameLogEntry> FrameLog { get; } = [];

    private string _statusMessage = "Ready";
    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    #endregion

    #region Commands

    public RelayCommand ConnectCommand { get; }
    public RelayCommand DisconnectCommand { get; }
    public RelayCommand RefreshPortsCommand { get; }
    public AsyncRelayCommand ExecuteMasterCommand { get; }
    public AsyncRelayCommand StartSlaveCommand { get; }
    public RelayCommand StopSlaveCommand { get; }
    public RelayCommand AddSlaveRegisterCommand { get; }
    public RelayCommand RemoveSlaveRegisterCommand { get; }
    public AsyncRelayCommand SendCustomFrameCommand { get; }
    public RelayCommand ClearLogCommand { get; }

    #endregion

    #region Methods

    private void RefreshPorts()
    {
        AvailableComPorts.Clear();
        foreach (var port in SerialPort.GetPortNames())
            AvailableComPorts.Add(port);
        if (AvailableComPorts.Count > 0)
            SelectedComPort = AvailableComPorts[0];
    }

    private ModbusProtocolType GetProtocolType() => SelectedProtocol switch
    {
        "RTU" => ModbusProtocolType.RTU,
        "ASCII" => ModbusProtocolType.ASCII,
        "TCP" => ModbusProtocolType.TCP,
        _ => ModbusProtocolType.TCP
    };

    private StopBits GetStopBits() => SelectedStopBits switch
    {
        "1" => StopBits.One,
        "1.5" => StopBits.OnePointFive,
        "2" => StopBits.Two,
        _ => StopBits.One
    };

    private Parity GetParity() => SelectedParity switch
    {
        "None" => Parity.None,
        "Odd" => Parity.Odd,
        "Even" => Parity.Even,
        "Mark" => Parity.Mark,
        "Space" => Parity.Space,
        _ => Parity.None
    };

    private void Connect()
    {
        try
        {
            if (IsTcpMode)
                _transport.ConnectTcp(IpAddress, TcpPort);
            else
                _transport.ConnectSerial(SelectedComPort, SelectedBaudRate, SelectedDataBits, GetStopBits(), GetParity());

            IsConnected = true;
            _failedHealthChecks = 0;
            _healthTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            _healthTimer.Tick += CheckConnectionHealth;
            _healthTimer.Start();
            StatusMessage = $"Connected via {SelectedProtocol}" +
                (IsTcpMode ? $" to {IpAddress}:{TcpPort}" : $" on {SelectedComPort}");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Connection failed: {ex.Message}";
            MessageBox.Show($"Connection failed:\n{ex.Message}", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Disconnect()
    {
        _healthTimer?.Stop();
        _transport.Disconnect();
        IsConnected = false;
        StatusMessage = "Disconnected";
    }

    private async Task ExecuteMaster()
    {
        try
        {
            StatusMessage = "Sending request...";
            ResponseInfo = "";
            RegisterResults.Clear();
            CoilResults.Clear();
            ShowRegisterResults = false;
            ShowCoilResults = false;

            var protocol = GetProtocolType();
            var functionCode = SelectedFunctionCode.Code;
            byte[] pduData;

            switch (functionCode)
            {
                case ModbusFunctionCode.ReadCoils:
                case ModbusFunctionCode.ReadDiscreteInputs:
                case ModbusFunctionCode.ReadHoldingRegisters:
                case ModbusFunctionCode.ReadInputRegisters:
                    pduData = ModbusProtocol.BuildReadRequestData(StartAddress, Quantity);
                    break;

                case ModbusFunctionCode.WriteSingleCoil:
                    pduData = ModbusProtocol.BuildWriteSingleCoilData(StartAddress, ParseSingleWriteValue() != 0);
                    break;

                case ModbusFunctionCode.WriteSingleRegister:
                    pduData = ModbusProtocol.BuildWriteSingleRegisterData(StartAddress, ParseSingleWriteValue());
                    break;

                case ModbusFunctionCode.WriteMultipleRegisters:
                    pduData = ModbusProtocol.BuildWriteMultipleRegistersData(StartAddress, ParseMultipleWriteValues());
                    break;

                case ModbusFunctionCode.WriteMultipleCoils:
                    pduData = ModbusProtocol.BuildWriteMultipleCoilsData(StartAddress,
                        ParseMultipleWriteValues().Select(v => v != 0).ToArray());
                    break;

                default:
                    throw new InvalidOperationException("Unsupported function code");
            }

            var frame = ModbusProtocol.BuildFrame(protocol, SlaveId, (byte)functionCode, pduData);

            var (response, commError) = await RunSafe(() => _transport.SendAndReceive(frame, protocol, TimeoutMs));
            if (commError != null)
            {
                HandleCommunicationError(commError);
                return;
            }

            var (_, respFc, respData) = ModbusProtocol.ParseResponse(protocol, response!);

            if ((respFc & 0x80) != 0)
            {
                var exCode = respData.Length > 0 ? respData[0] : (byte)0;
                ResponseInfo = $"? Modbus Exception: {ModbusProtocol.GetExceptionMessage(exCode)}";
                StatusMessage = "Request completed with error";
                return;
            }

            switch (functionCode)
            {
                case ModbusFunctionCode.ReadHoldingRegisters:
                case ModbusFunctionCode.ReadInputRegisters:
                {
                    var registers = ModbusProtocol.ExtractRegisters(respData, CurrentByteOrder);
                    for (int i = 0; i < registers.Length; i++)
                        RegisterResults.Add(new RegisterResult { Address = (ushort)(StartAddress + i), RawValue = registers[i] });
                    ShowRegisterResults = true;
                    ResponseInfo = $"? Read {registers.Length} registers starting at address {StartAddress}";
                    break;
                }
                case ModbusFunctionCode.ReadCoils:
                case ModbusFunctionCode.ReadDiscreteInputs:
                {
                    var coils = ModbusProtocol.ExtractCoils(respData, Quantity);
                    for (int i = 0; i < coils.Length; i++)
                        CoilResults.Add(new CoilResult { Address = (ushort)(StartAddress + i), Value = coils[i] });
                    ShowCoilResults = true;
                    ResponseInfo = $"? Read {coils.Length} coils/inputs starting at address {StartAddress}";
                    break;
                }
                case ModbusFunctionCode.WriteSingleCoil:
                case ModbusFunctionCode.WriteSingleRegister:
                    ResponseInfo = $"? Write successful at address {StartAddress}";
                    break;
                case ModbusFunctionCode.WriteMultipleRegisters:
                case ModbusFunctionCode.WriteMultipleCoils:
                    ResponseInfo = $"? Write successful: values starting at address {StartAddress}";
                    break;
            }
            StatusMessage = "Request completed successfully";
        }
        catch (Exception ex)
        {
            ResponseInfo = $"\u2717 Error: {ex.Message}";
            StatusMessage = $"Error: {ex.Message}";
        }
    }

    private ushort ParseSingleWriteValue()
    {
        var text = WriteValues.Trim();
        if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            return Convert.ToUInt16(text[2..], 16);
        return ushort.Parse(text);
    }

    private ushort[] ParseMultipleWriteValues()
    {
        var parts = WriteValues.Split([',', ' ', ';'], StringSplitOptions.RemoveEmptyEntries);
        var values = new ushort[parts.Length];
        for (int i = 0; i < parts.Length; i++)
        {
            var text = parts[i].Trim();
            if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                values[i] = Convert.ToUInt16(text[2..], 16);
            else
                values[i] = ushort.Parse(text);
        }
        return values;
    }

    private async Task StartSlave()
    {
        try
        {
            _slaveCts = new CancellationTokenSource();
            IsSlaveRunning = true;

            var protocol = GetProtocolType();

            if (protocol == ModbusProtocolType.TCP)
            {
                SlaveStatus = $"Listening on TCP port {TcpPort}...";
                await _transport.StartTcpSlaveAsync(TcpPort, HandleSlaveRequest,
                    data => AddLog("RX ?", data),
                    data => AddLog("TX ?", data),
                    _slaveCts.Token);
            }
            else
            {
                SlaveStatus = $"Listening on {SelectedComPort} ({SelectedProtocol})...";
                await _transport.StartSerialSlaveAsync(SelectedComPort, SelectedBaudRate, SelectedDataBits,
                    GetStopBits(), GetParity(), protocol, SlaveId, HandleSlaveRequest,
                    data => AddLog("RX ?", data),
                    data => AddLog("TX ?", data),
                    _slaveCts.Token);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            _dispatcher.Invoke(() => SlaveStatus = $"Error: {ex.Message}");
        }
        finally
        {
            _dispatcher.Invoke(() =>
            {
                IsSlaveRunning = false;
                SlaveStatus = "Stopped";
            });
        }
    }

    private byte[]? HandleSlaveRequest(byte functionCode, byte[] requestData)
    {
        if (requestData.Length < 2) return null;

        var startAddress = (ushort)((requestData[0] << 8) | requestData[1]);

        List<SlaveDataEntry> entries = [];
        _dispatcher.Invoke(() => entries = [.. SlaveRegisters]);

        switch ((ModbusFunctionCode)functionCode)
        {
            case ModbusFunctionCode.ReadHoldingRegisters:
            case ModbusFunctionCode.ReadInputRegisters:
            {
                if (requestData.Length < 4) return null;
                var quantity = (ushort)((requestData[2] << 8) | requestData[3]);
                var byteCount = (byte)(quantity * 2);
                var response = new byte[1 + byteCount];
                response[0] = byteCount;
                for (int i = 0; i < quantity; i++)
                {
                    var addr = (ushort)(startAddress + i);
                    var entry = entries.FirstOrDefault(e => e.Address == addr);
                    var value = entry?.Value ?? 0;
                    response[1 + i * 2] = (byte)((value >> 8) & 0xFF);
                    response[2 + i * 2] = (byte)(value & 0xFF);
                }
                return response;
            }
            case ModbusFunctionCode.ReadCoils:
            case ModbusFunctionCode.ReadDiscreteInputs:
            {
                if (requestData.Length < 4) return null;
                var quantity = (ushort)((requestData[2] << 8) | requestData[3]);
                var byteCount = (byte)((quantity + 7) / 8);
                var response = new byte[1 + byteCount];
                response[0] = byteCount;
                for (int i = 0; i < quantity; i++)
                {
                    var addr = (ushort)(startAddress + i);
                    var entry = entries.FirstOrDefault(e => e.Address == addr);
                    if (entry != null && entry.Value != 0)
                        response[1 + i / 8] |= (byte)(1 << (i % 8));
                }
                return response;
            }
            case ModbusFunctionCode.WriteSingleRegister:
            {
                if (requestData.Length < 4) return null;
                var value = (ushort)((requestData[2] << 8) | requestData[3]);
                _dispatcher.Invoke(() =>
                {
                    var entry = SlaveRegisters.FirstOrDefault(e => e.Address == startAddress);
                    if (entry != null) entry.Value = value;
                    else SlaveRegisters.Add(new SlaveDataEntry { Address = startAddress, Value = value });
                });
                return requestData;
            }
            case ModbusFunctionCode.WriteSingleCoil:
            {
                if (requestData.Length < 4) return null;
                var value = (ushort)(requestData[2] == 0xFF ? 1 : 0);
                _dispatcher.Invoke(() =>
                {
                    var entry = SlaveRegisters.FirstOrDefault(e => e.Address == startAddress);
                    if (entry != null) entry.Value = value;
                    else SlaveRegisters.Add(new SlaveDataEntry { Address = startAddress, Value = value });
                });
                return requestData;
            }
            case ModbusFunctionCode.WriteMultipleRegisters:
            {
                if (requestData.Length < 5) return null;
                var quantity = (ushort)((requestData[2] << 8) | requestData[3]);
                _dispatcher.Invoke(() =>
                {
                    for (int i = 0; i < quantity && (5 + i * 2 + 1) < requestData.Length; i++)
                    {
                        var addr = (ushort)(startAddress + i);
                        var val = (ushort)((requestData[5 + i * 2] << 8) | requestData[6 + i * 2]);
                        var entry = SlaveRegisters.FirstOrDefault(e => e.Address == addr);
                        if (entry != null) entry.Value = val;
                        else SlaveRegisters.Add(new SlaveDataEntry { Address = addr, Value = val });
                    }
                });
                return requestData[..4];
            }
            default:
                return null;
        }
    }

    private void StopSlave()
    {
        _slaveCts?.Cancel();
        _transport.StopSlave();
    }

    private void AddSlaveRegister()
    {
        ushort nextAddr = SlaveRegisters.Count > 0
            ? (ushort)(SlaveRegisters.Max(r => r.Address) + 1)
            : (ushort)0;
        SlaveRegisters.Add(new SlaveDataEntry { Address = nextAddr, Value = 0 });
    }

    private void RemoveSlaveRegister()
    {
        if (SlaveRegisters.Count > 0)
            SlaveRegisters.RemoveAt(SlaveRegisters.Count - 1);
    }

    private async Task SendCustomFrame()
    {
        try
        {
            StatusMessage = "Sending custom frame...";
            CustomResponseHex = "";

            var hex = CustomFrameHex.Replace(" ", "").Replace("-", "");
            if (hex.Length % 2 != 0)
                throw new FormatException("Hex string must have even number of characters");

            var bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
                bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);

            byte[] frame;
            if (AutoAddChecksum)
            {
                var protocol = GetProtocolType();
                if (protocol == ModbusProtocolType.RTU)
                {
                    var crc = ModbusProtocol.CalculateCRC16(bytes, 0, bytes.Length);
                    frame = new byte[bytes.Length + 2];
                    Array.Copy(bytes, frame, bytes.Length);
                    frame[^2] = (byte)(crc & 0xFF);
                    frame[^1] = (byte)((crc >> 8) & 0xFF);
                }
                else if (protocol == ModbusProtocolType.ASCII)
                {
                    var lrc = ModbusProtocol.CalculateLRC(bytes, 0, bytes.Length);
                    var hexStr = ":" + BitConverter.ToString(bytes).Replace("-", "") + lrc.ToString("X2") + "\r\n";
                    frame = Encoding.ASCII.GetBytes(hexStr);
                }
                else
                {
                    frame = bytes;
                }
            }
            else
            {
                frame = bytes;
            }

            if (WaitForResponse)
            {
                var (response, commError) = await RunSafe(() => _transport.SendAndReceive(frame, GetProtocolType(), TimeoutMs));
                if (commError != null)
                {
                    HandleCommunicationError(commError, isCustomFrame: true);
                    return;
                }
                CustomResponseHex = BitConverter.ToString(response!).Replace("-", " ");
                StatusMessage = "Custom frame sent, response received";
            }
            else
            {
                var sendError = await RunSafe(() => _transport.Send(frame));
                if (sendError != null)
                {
                    HandleCommunicationError(sendError, isCustomFrame: true);
                    return;
                }
                StatusMessage = "Custom frame sent";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error: {ex.Message}";
            CustomResponseHex = $"Error: {ex.Message}";
        }
    }

    private void AddLog(string direction, byte[] data)
    {
        var entry = new FrameLogEntry(DateTime.Now, direction, data);
        _dispatcher.BeginInvoke(() => FrameLog.Add(entry));
    }

    private static async Task<(T? result, Exception? error)> RunSafe<T>(Func<T> action)
    {
        T? result = default;
        Exception? error = null;
        await Task.Run(() =>
        {
            try { result = action(); }
            catch (Exception ex) { error = ex; }
        });
        return (result, error);
    }

    private static async Task<Exception?> RunSafe(Action action)
    {
        Exception? error = null;
        await Task.Run(() =>
        {
            try { action(); }
            catch (Exception ex) { error = ex; }
        });
        return error;
    }

    private void HandleCommunicationError(Exception ex, bool isCustomFrame = false)
    {
        // Connection lost — auto-disconnect
        if (!_transport.IsAlive && IsConnected)
        {
            AutoDisconnect($"Connection lost: {ex.Message}");
            return;
        }

        if (ex is InvalidOperationException &&
            (ex.Message.Contains("closed", StringComparison.OrdinalIgnoreCase) ||
             ex.Message.Contains("lost", StringComparison.OrdinalIgnoreCase) ||
             ex.Message.Contains("disposed", StringComparison.OrdinalIgnoreCase) ||
             ex is ObjectDisposedException))
        {
            AutoDisconnect($"Connection lost during operation:\n{ex.Message}");
            return;
        }

        string title, popupMessage, hint;

        if (ex is TimeoutException)
        {
            title = "Communication Timeout";
            popupMessage = "No response received from the device.";
            hint = "\u2022 Check that the device is powered on and connected\n" +
                   "\u2022 Verify the IP address/COM port is correct\n" +
                   "\u2022 Confirm the Slave ID matches the device\n" +
                   "\u2022 Try increasing the timeout value\n" +
                   "\u2022 Check your network/serial cable connection";

            StatusMessage = "Timeout - no response received";
            if (isCustomFrame)
                CustomResponseHex = "No response received - timeout";
            else
                ResponseInfo = "\u2717 No response - device did not reply within the timeout period";
        }
        else
        {
            title = "Communication Error";
            popupMessage = $"An error occurred:\n{ex.Message}";
            hint = "\u2022 Check your connection settings\n" +
                   "\u2022 Verify the device is responding\n" +
                   "\u2022 Try disconnecting and reconnecting";

            StatusMessage = $"Error: {ex.Message}";
            if (isCustomFrame)
                CustomResponseHex = $"Error: {ex.Message}";
            else
                ResponseInfo = $"\u2717 Error: {ex.Message}";
        }

        MessageBox.Show(
            $"{popupMessage}\n\nSuggestions:\n{hint}",
            title,
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private void CheckConnectionHealth(object? sender, EventArgs e)
    {
        if (!IsConnected) { _healthTimer?.Stop(); return; }

        if (!_transport.IsAlive)
        {
            _failedHealthChecks++;
            if (_failedHealthChecks >= 6) // 30 seconds (6 × 5s)
            {
                AutoDisconnect(
                    "The device has not responded for over 30 seconds.\n" +
                    "The connection appears to have been lost.");
            }
        }
        else
        {
            _failedHealthChecks = 0;
        }
    }

    private void AutoDisconnect(string reason)
    {
        _healthTimer?.Stop();
        _transport.Disconnect();
        IsConnected = false;
        StatusMessage = "Disconnected - connection lost";
        ResponseInfo = "";
        MessageBox.Show(
            $"{reason}\n\n" +
            "The application has been disconnected automatically.\n\n" +
            "To resume:\n" +
            "\u2022 Check that the device is powered on and connected\n" +
            "\u2022 Verify your cable/network connection\n" +
            "\u2022 Click Connect to reconnect",
            "Connection Lost",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private void HandleCommandError(Exception ex)
    {
        // If the transport is dead, auto-disconnect
        if (!_transport.IsAlive && IsConnected)
        {
            AutoDisconnect($"Connection lost: {ex.Message}");
            return;
        }

        var (title, message, hint) = ex switch
        {
            TimeoutException => (
                "Communication Timeout",
                "No response received from the device within the timeout period.",
                "\u2022 Check that the device is powered on and connected\n" +
                "\u2022 Verify the IP address/COM port and Slave ID\n" +
                "\u2022 Try increasing the timeout value"),
            InvalidOperationException when ex.Message.Contains("closed", StringComparison.OrdinalIgnoreCase) => (
                "Connection Lost",
                "The connection to the device was lost.",
                "\u2022 Check your network/serial cable\n" +
                "\u2022 Click Disconnect then Connect to reconnect"),
            _ => (
                "Communication Error",
                ex.Message,
                "\u2022 Check your connection settings\n" +
                "\u2022 Verify the device is responding\n" +
                "\u2022 Try disconnecting and reconnecting")
        };

        StatusMessage = $"Error: {message}";
        MessageBox.Show(
            $"{message}\n\nSuggestions:\n{hint}",
            title,
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    public void Dispose()
    {
        _healthTimer?.Stop();
        _slaveCts?.Cancel();
        _transport.Dispose();
    }

    #endregion
}
