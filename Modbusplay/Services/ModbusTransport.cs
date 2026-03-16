using System.IO;
using System.IO.Ports;
using System.Net;
using System.Net.Sockets;
using Modbusplay.Models;

namespace Modbusplay.Services;

public class ModbusTransport : IDisposable
{
    private SerialPort? _serialPort;
    private TcpClient? _tcpClient;
    private NetworkStream? _tcpStream;
    private TcpListener? _listener;
    private TcpClient? _slaveClient;

    public bool IsConnected { get; private set; }

    public event Action<byte[]>? DataSent;
    public event Action<byte[]>? DataReceived;
    public event Action<string>? ConnectionLost;

    public bool IsAlive
    {
        get
        {
            try
            {
                if (_serialPort != null) return _serialPort.IsOpen;
                if (_tcpClient != null) return _tcpClient.Connected;
                return false;
            }
            catch { return false; }
        }
    }

    private void MarkConnectionLost(string reason)
    {
        IsConnected = false;
        ConnectionLost?.Invoke(reason);
    }

    public void ConnectSerial(string portName, int baudRate, int dataBits, StopBits stopBits, Parity parity)
    {
        Disconnect();
        _serialPort = new SerialPort(portName, baudRate, parity, dataBits, stopBits);
        _serialPort.Open();
        IsConnected = true;
    }

    public void ConnectTcp(string ipAddress, int port)
    {
        Disconnect();
        _tcpClient = new TcpClient();
        _tcpClient.Connect(ipAddress, port);
        _tcpStream = _tcpClient.GetStream();
        IsConnected = true;
    }

    public void Disconnect()
    {
        try
        {
            _serialPort?.Close();
            _serialPort?.Dispose();
            _tcpStream?.Dispose();
            _tcpClient?.Dispose();
        }
        catch { }

        _serialPort = null;
        _tcpClient = null;
        _tcpStream = null;
        IsConnected = false;
    }

    public byte[] SendAndReceive(byte[] frame, ModbusProtocolType protocol, int timeoutMs)
    {
        Send(frame);
        return Receive(protocol, timeoutMs);
    }

    public void Send(byte[] frame)
    {
        if (!IsConnected)
            throw new InvalidOperationException("Not connected");

        try
        {
            if (_serialPort != null)
                _serialPort.Write(frame, 0, frame.Length);
            else if (_tcpStream != null)
                _tcpStream.Write(frame, 0, frame.Length);

            DataSent?.Invoke(frame);
        }
        catch (InvalidOperationException ex)
        {
            MarkConnectionLost($"Connection lost: {ex.Message}");
            throw;
        }
        catch (IOException ex)
        {
            if (!IsAlive) MarkConnectionLost($"Connection lost: {ex.Message}");
            throw;
        }
    }

    public byte[] Receive(ModbusProtocolType protocol, int timeoutMs)
    {
        if (!IsConnected)
            throw new InvalidOperationException("Not connected");

        try
        {
            var result = protocol switch
            {
                ModbusProtocolType.RTU => ReceiveRtu(timeoutMs),
                ModbusProtocolType.ASCII => ReceiveAscii(timeoutMs),
                ModbusProtocolType.TCP => ReceiveTcp(timeoutMs),
                _ => throw new ArgumentOutOfRangeException(nameof(protocol))
            };

            DataReceived?.Invoke(result);
            return result;
        }
        catch (InvalidOperationException ex) when (
            ex is ObjectDisposedException ||
            ex.Message.Contains("closed", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("not open", StringComparison.OrdinalIgnoreCase))
        {
            MarkConnectionLost($"Connection lost: {ex.Message}");
            throw;
        }
    }

    private byte[] ReceiveRtu(int timeoutMs)
    {
        if (_serialPort != null)
        {
            var buffer = new List<byte>();
            _serialPort.ReadTimeout = timeoutMs;
            try
            {
                buffer.Add((byte)_serialPort.ReadByte());
                _serialPort.ReadTimeout = 50;
                while (true)
                {
                    try { buffer.Add((byte)_serialPort.ReadByte()); }
                    catch (TimeoutException) { break; }
                }
            }
            catch (TimeoutException)
            {
                throw new TimeoutException("No response received within timeout period");
            }
            catch (IOException ex)
            {
                throw new TimeoutException($"Serial communication error: {ex.Message}", ex);
            }
            return [.. buffer];
        }

        return ReceiveTcp(timeoutMs);
    }

    private byte[] ReceiveAscii(int timeoutMs)
    {
        if (_serialPort == null)
            throw new InvalidOperationException("Serial port not connected");

        var buffer = new List<byte>();
        _serialPort.ReadTimeout = timeoutMs;
        try
        {
            while (true)
            {
                var b = (byte)_serialPort.ReadByte();
                buffer.Add(b);
                if (buffer.Count >= 2 && buffer[^2] == 0x0D && buffer[^1] == 0x0A)
                    break;
            }
        }
        catch (TimeoutException)
        {
            if (buffer.Count == 0)
                throw new TimeoutException("No response received within timeout period");
        }
        catch (IOException ex)
        {
            throw new TimeoutException($"Serial communication error: {ex.Message}", ex);
        }
        return [.. buffer];
    }

    private byte[] ReceiveTcp(int timeoutMs)
    {
        if (_tcpStream == null)
            throw new InvalidOperationException("TCP not connected");

        _tcpStream.ReadTimeout = timeoutMs;

        try
        {
            var header = new byte[6];
            int totalRead = 0;
            while (totalRead < 6)
            {
                int read = _tcpStream.Read(header, totalRead, 6 - totalRead);
                if (read == 0) throw new InvalidOperationException("Connection closed by remote host");
                totalRead += read;
            }

            int length = (header[4] << 8) | header[5];

            var pdu = new byte[length];
            totalRead = 0;
            while (totalRead < length)
            {
                int read = _tcpStream.Read(pdu, totalRead, length - totalRead);
                if (read == 0) throw new InvalidOperationException("Connection closed by remote host");
                totalRead += read;
            }

            var frame = new byte[6 + length];
            Array.Copy(header, 0, frame, 0, 6);
            Array.Copy(pdu, 0, frame, 6, length);
            return frame;
        }
        catch (IOException ex) when (ex.InnerException is SocketException { SocketErrorCode: SocketError.TimedOut })
        {
            throw new TimeoutException("No response received within timeout period", ex);
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException($"TCP communication error: {ex.Message}", ex);
        }
    }

    public async Task StartTcpSlaveAsync(int port, Func<byte, byte[], byte[]?> requestHandler, Action<byte[]> onReceived, Action<byte[]> onSent, CancellationToken ct)
    {
        _listener = new TcpListener(IPAddress.Any, port);
        _listener.Start();

        try
        {
            while (!ct.IsCancellationRequested)
            {
                _slaveClient = await _listener.AcceptTcpClientAsync(ct);
                var stream = _slaveClient.GetStream();

                try
                {
                    while (!ct.IsCancellationRequested && _slaveClient.Connected)
                    {
                        var frame = await ReadTcpFrameAsync(stream, ct);
                        if (frame == null) break;

                        onReceived(frame);

                        var unitId = frame[6];
                        var functionCode = frame[7];
                        var data = new byte[frame.Length - 8];
                        Array.Copy(frame, 8, data, 0, data.Length);

                        var responseData = requestHandler(functionCode, data);
                        if (responseData != null)
                        {
                            var response = ModbusProtocol.BuildTcpFrame(unitId, functionCode, responseData);
                            response[0] = frame[0];
                            response[1] = frame[1];

                            stream.Write(response, 0, response.Length);
                            onSent(response);
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch { }
                finally
                {
                    _slaveClient?.Dispose();
                }
            }
        }
        finally
        {
            _listener.Stop();
        }
    }

    public async Task StartSerialSlaveAsync(string portName, int baudRate, int dataBits, StopBits stopBits, Parity parity,
        ModbusProtocolType protocol, byte slaveId, Func<byte, byte[], byte[]?> requestHandler, Action<byte[]> onReceived, Action<byte[]> onSent, CancellationToken ct)
    {
        var slaveSerial = new SerialPort(portName, baudRate, parity, dataBits, stopBits);
        slaveSerial.Open();

        try
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    byte[] frame = protocol == ModbusProtocolType.ASCII
                        ? await Task.Run(() => ReceiveAsciiSlave(slaveSerial, ct), ct)
                        : await Task.Run(() => ReceiveRtuSlave(slaveSerial, ct), ct);

                    if (frame.Length == 0) continue;

                    onReceived(frame);

                    var parsed = ModbusProtocol.ParseResponse(protocol, frame);
                    if (parsed.slaveId != slaveId) continue;

                    var responseData = requestHandler(parsed.functionCode, parsed.data);
                    if (responseData != null)
                    {
                        var response = ModbusProtocol.BuildFrame(protocol, slaveId, parsed.functionCode, responseData);
                        slaveSerial.Write(response, 0, response.Length);
                        onSent(response);
                    }
                }
                catch (TimeoutException) { }
                catch (OperationCanceledException) { throw; }
                catch { }
            }
        }
        finally
        {
            slaveSerial.Close();
            slaveSerial.Dispose();
        }
    }

    private static byte[] ReceiveRtuSlave(SerialPort serial, CancellationToken ct)
    {
        serial.ReadTimeout = 500;
        var buffer = new List<byte>();

        while (!ct.IsCancellationRequested)
        {
            try
            {
                buffer.Add((byte)serial.ReadByte());
                serial.ReadTimeout = 50;
                while (true)
                {
                    try { buffer.Add((byte)serial.ReadByte()); }
                    catch (TimeoutException) { break; }
                }
                return [.. buffer];
            }
            catch (TimeoutException) { }
        }
        return [];
    }

    private static byte[] ReceiveAsciiSlave(SerialPort serial, CancellationToken ct)
    {
        serial.ReadTimeout = 500;
        var buffer = new List<byte>();

        while (!ct.IsCancellationRequested)
        {
            try
            {
                var b = (byte)serial.ReadByte();
                buffer.Add(b);
                if (buffer.Count >= 2 && buffer[^2] == 0x0D && buffer[^1] == 0x0A)
                    return [.. buffer];
            }
            catch (TimeoutException)
            {
                buffer.Clear();
            }
        }
        return [];
    }

    private static async Task<byte[]?> ReadTcpFrameAsync(NetworkStream stream, CancellationToken ct)
    {
        var header = new byte[6];
        int totalRead = 0;
        while (totalRead < 6)
        {
            int read = await stream.ReadAsync(header.AsMemory(totalRead, 6 - totalRead), ct);
            if (read == 0) return null;
            totalRead += read;
        }

        int length = (header[4] << 8) | header[5];
        var pdu = new byte[length];
        totalRead = 0;
        while (totalRead < length)
        {
            int read = await stream.ReadAsync(pdu.AsMemory(totalRead, length - totalRead), ct);
            if (read == 0) return null;
            totalRead += read;
        }

        var frame = new byte[6 + length];
        Array.Copy(header, 0, frame, 0, 6);
        Array.Copy(pdu, 0, frame, 6, length);
        return frame;
    }

    public void StopSlave()
    {
        try { _listener?.Stop(); } catch { }
        try { _slaveClient?.Dispose(); } catch { }
    }

    public void Dispose()
    {
        Disconnect();
        StopSlave();
    }
}
