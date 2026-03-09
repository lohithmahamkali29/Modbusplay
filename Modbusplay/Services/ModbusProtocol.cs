using System.Text;
using Modbusplay.Models;

namespace Modbusplay.Services;

public static class ModbusProtocol
{
    private static ushort _transactionId;

    public static ushort CalculateCRC16(byte[] data, int offset, int count)
    {
        ushort crc = 0xFFFF;
        for (int i = offset; i < offset + count; i++)
        {
            crc ^= data[i];
            for (int j = 0; j < 8; j++)
            {
                if ((crc & 0x0001) != 0)
                    crc = (ushort)((crc >> 1) ^ 0xA001);
                else
                    crc >>= 1;
            }
        }
        return crc;
    }

    public static byte CalculateLRC(byte[] data, int offset, int count)
    {
        byte lrc = 0;
        for (int i = offset; i < offset + count; i++)
            lrc += data[i];
        return (byte)(-(sbyte)lrc);
    }

    public static byte[] BuildFrame(ModbusProtocolType protocol, byte slaveId, byte functionCode, byte[] pduData)
    {
        return protocol switch
        {
            ModbusProtocolType.RTU => BuildRtuFrame(slaveId, functionCode, pduData),
            ModbusProtocolType.ASCII => BuildAsciiFrame(slaveId, functionCode, pduData),
            ModbusProtocolType.TCP => BuildTcpFrame(slaveId, functionCode, pduData),
            _ => throw new ArgumentOutOfRangeException(nameof(protocol))
        };
    }

    public static byte[] BuildRtuFrame(byte slaveId, byte functionCode, byte[] pduData)
    {
        var frame = new byte[2 + pduData.Length + 2];
        frame[0] = slaveId;
        frame[1] = functionCode;
        Array.Copy(pduData, 0, frame, 2, pduData.Length);
        var crc = CalculateCRC16(frame, 0, frame.Length - 2);
        frame[^2] = (byte)(crc & 0xFF);
        frame[^1] = (byte)((crc >> 8) & 0xFF);
        return frame;
    }

    public static byte[] BuildAsciiFrame(byte slaveId, byte functionCode, byte[] pduData)
    {
        var dataBytes = new byte[2 + pduData.Length];
        dataBytes[0] = slaveId;
        dataBytes[1] = functionCode;
        Array.Copy(pduData, 0, dataBytes, 2, pduData.Length);

        var lrc = CalculateLRC(dataBytes, 0, dataBytes.Length);
        var hex = BitConverter.ToString(dataBytes).Replace("-", "") + lrc.ToString("X2");
        var ascii = $":{hex}\r\n";
        return Encoding.ASCII.GetBytes(ascii);
    }

    public static byte[] BuildTcpFrame(byte unitId, byte functionCode, byte[] pduData)
    {
        _transactionId++;
        var length = (ushort)(1 + 1 + pduData.Length);
        var frame = new byte[6 + length];

        frame[0] = (byte)((_transactionId >> 8) & 0xFF);
        frame[1] = (byte)(_transactionId & 0xFF);
        frame[2] = 0;
        frame[3] = 0;
        frame[4] = (byte)((length >> 8) & 0xFF);
        frame[5] = (byte)(length & 0xFF);
        frame[6] = unitId;
        frame[7] = functionCode;
        Array.Copy(pduData, 0, frame, 8, pduData.Length);

        return frame;
    }

    public static (byte slaveId, byte functionCode, byte[] data) ParseResponse(ModbusProtocolType protocol, byte[] frame)
    {
        return protocol switch
        {
            ModbusProtocolType.RTU => ParseRtuResponse(frame),
            ModbusProtocolType.ASCII => ParseAsciiResponse(frame),
            ModbusProtocolType.TCP => ParseTcpResponse(frame),
            _ => throw new ArgumentOutOfRangeException(nameof(protocol))
        };
    }

    public static (byte slaveId, byte functionCode, byte[] data) ParseRtuResponse(byte[] frame)
    {
        if (frame.Length < 4)
            throw new InvalidOperationException("RTU frame too short");

        var crc = CalculateCRC16(frame, 0, frame.Length - 2);
        var receivedCrc = (ushort)(frame[^2] | (frame[^1] << 8));
        if (crc != receivedCrc)
            throw new InvalidOperationException($"CRC mismatch: expected 0x{crc:X4}, received 0x{receivedCrc:X4}");

        var slaveId = frame[0];
        var functionCode = frame[1];
        var data = new byte[frame.Length - 4];
        Array.Copy(frame, 2, data, 0, data.Length);
        return (slaveId, functionCode, data);
    }

    public static (byte slaveId, byte functionCode, byte[] data) ParseAsciiResponse(byte[] frame)
    {
        var ascii = Encoding.ASCII.GetString(frame).Trim();
        if (!ascii.StartsWith(':'))
            throw new InvalidOperationException("Invalid ASCII frame: missing start character ':'");

        ascii = ascii[1..];
        if (ascii.EndsWith("\r\n"))
            ascii = ascii[..^2];

        var bytes = new byte[ascii.Length / 2];
        for (int i = 0; i < bytes.Length; i++)
            bytes[i] = Convert.ToByte(ascii.Substring(i * 2, 2), 16);

        var lrc = CalculateLRC(bytes, 0, bytes.Length - 1);
        if (lrc != bytes[^1])
            throw new InvalidOperationException($"LRC mismatch: expected 0x{lrc:X2}, received 0x{bytes[^1]:X2}");

        var slaveId = bytes[0];
        var functionCode = bytes[1];
        var data = new byte[bytes.Length - 3];
        Array.Copy(bytes, 2, data, 0, data.Length);
        return (slaveId, functionCode, data);
    }

    public static (byte slaveId, byte functionCode, byte[] data) ParseTcpResponse(byte[] frame)
    {
        if (frame.Length < 9)
            throw new InvalidOperationException("TCP frame too short");

        var unitId = frame[6];
        var functionCode = frame[7];
        var data = new byte[frame.Length - 8];
        Array.Copy(frame, 8, data, 0, data.Length);
        return (unitId, functionCode, data);
    }

    public static byte[] BuildReadRequestData(ushort startAddress, ushort quantity)
    {
        return
        [
            (byte)((startAddress >> 8) & 0xFF),
            (byte)(startAddress & 0xFF),
            (byte)((quantity >> 8) & 0xFF),
            (byte)(quantity & 0xFF)
        ];
    }

    public static byte[] BuildWriteSingleCoilData(ushort address, bool value)
    {
        return
        [
            (byte)((address >> 8) & 0xFF),
            (byte)(address & 0xFF),
            (byte)(value ? 0xFF : 0x00),
            0x00
        ];
    }

    public static byte[] BuildWriteSingleRegisterData(ushort address, ushort value)
    {
        return
        [
            (byte)((address >> 8) & 0xFF),
            (byte)(address & 0xFF),
            (byte)((value >> 8) & 0xFF),
            (byte)(value & 0xFF)
        ];
    }

    public static byte[] BuildWriteMultipleRegistersData(ushort startAddress, ushort[] values)
    {
        var byteCount = (byte)(values.Length * 2);
        var data = new byte[4 + 1 + byteCount];
        data[0] = (byte)((startAddress >> 8) & 0xFF);
        data[1] = (byte)(startAddress & 0xFF);
        data[2] = (byte)((values.Length >> 8) & 0xFF);
        data[3] = (byte)(values.Length & 0xFF);
        data[4] = byteCount;
        for (int i = 0; i < values.Length; i++)
        {
            data[5 + i * 2] = (byte)((values[i] >> 8) & 0xFF);
            data[6 + i * 2] = (byte)(values[i] & 0xFF);
        }
        return data;
    }

    public static byte[] BuildWriteMultipleCoilsData(ushort startAddress, bool[] values)
    {
        var byteCount = (byte)((values.Length + 7) / 8);
        var data = new byte[4 + 1 + byteCount];
        data[0] = (byte)((startAddress >> 8) & 0xFF);
        data[1] = (byte)(startAddress & 0xFF);
        data[2] = (byte)((values.Length >> 8) & 0xFF);
        data[3] = (byte)(values.Length & 0xFF);
        data[4] = byteCount;
        for (int i = 0; i < values.Length; i++)
        {
            if (values[i])
                data[5 + i / 8] |= (byte)(1 << (i % 8));
        }
        return data;
    }

    public static ushort[] ExtractRegisters(byte[] responseData, ByteOrder byteOrder)
    {
        if (responseData.Length < 1) return [];

        var byteCount = responseData[0];
        var registers = new ushort[byteCount / 2];

        for (int i = 0; i < registers.Length; i++)
        {
            if (byteOrder == ByteOrder.MSDFirst)
                registers[i] = (ushort)((responseData[1 + i * 2] << 8) | responseData[2 + i * 2]);
            else
                registers[i] = (ushort)((responseData[2 + i * 2] << 8) | responseData[1 + i * 2]);
        }

        return registers;
    }

    public static bool[] ExtractCoils(byte[] responseData, int quantity)
    {
        if (responseData.Length < 1) return [];

        var coils = new bool[quantity];
        for (int i = 0; i < quantity; i++)
        {
            coils[i] = (responseData[1 + i / 8] & (1 << (i % 8))) != 0;
        }

        return coils;
    }

    public static string GetExceptionMessage(byte exceptionCode)
    {
        return exceptionCode switch
        {
            0x01 => "Illegal Function",
            0x02 => "Illegal Data Address",
            0x03 => "Illegal Data Value",
            0x04 => "Server Device Failure",
            0x05 => "Acknowledge",
            0x06 => "Server Device Busy",
            0x08 => "Memory Parity Error",
            0x0A => "Gateway Path Unavailable",
            0x0B => "Gateway Target Device Failed to Respond",
            _ => $"Unknown Exception (0x{exceptionCode:X2})"
        };
    }
}
