using Modbusplay.Helpers;

namespace Modbusplay.Models;

public enum ModbusProtocolType
{
    RTU,
    ASCII,
    TCP
}

public enum ModbusFunctionCode : byte
{
    ReadCoils = 0x01,
    ReadDiscreteInputs = 0x02,
    ReadHoldingRegisters = 0x03,
    ReadInputRegisters = 0x04,
    WriteSingleCoil = 0x05,
    WriteSingleRegister = 0x06,
    WriteMultipleCoils = 0x0F,
    WriteMultipleRegisters = 0x10
}

public enum ByteOrder
{
    MSDFirst,
    LSDFirst
}

public class FunctionCodeItem(ModbusFunctionCode code, string displayName)
{
    public ModbusFunctionCode Code { get; } = code;
    public string DisplayName { get; } = displayName;

    public bool IsWrite => Code is ModbusFunctionCode.WriteSingleCoil
        or ModbusFunctionCode.WriteSingleRegister
        or ModbusFunctionCode.WriteMultipleCoils
        or ModbusFunctionCode.WriteMultipleRegisters;

    public bool IsSingleWrite => Code is ModbusFunctionCode.WriteSingleCoil
        or ModbusFunctionCode.WriteSingleRegister;

    public bool IsMultipleWrite => Code is ModbusFunctionCode.WriteMultipleCoils
        or ModbusFunctionCode.WriteMultipleRegisters;

    public bool IsCoilFunction => Code is ModbusFunctionCode.ReadCoils
        or ModbusFunctionCode.ReadDiscreteInputs
        or ModbusFunctionCode.WriteSingleCoil
        or ModbusFunctionCode.WriteMultipleCoils;

    public override string ToString() => DisplayName;
}

public class RegisterResult
{
    public required ushort Address { get; init; }
    public required ushort RawValue { get; init; }
    public string HexValue => $"0x{RawValue:X4}";
    public string DecimalValue => RawValue.ToString();
    public string SignedValue => ((short)RawValue).ToString();
    public string BinaryValue => Convert.ToString(RawValue, 2).PadLeft(16, '0');
}

public class CoilResult
{
    public required ushort Address { get; init; }
    public required bool Value { get; init; }
    public string Status => Value ? "ON" : "OFF";
    public string HexValue => Value ? "FF00" : "0000";
}

public class FrameLogEntry(DateTime timestamp, string direction, byte[] data, string description = "")
{
    public DateTime Timestamp { get; } = timestamp;
    public string Direction { get; } = direction;
    public string HexData { get; } = BitConverter.ToString(data).Replace("-", " ");
    public string Description { get; } = description;
    public string Display => $"[{Timestamp:HH:mm:ss.fff}] {Direction} {HexData}{(string.IsNullOrEmpty(Description) ? "" : $"  {Description}")}";
}

public class SlaveDataEntry : ObservableObject
{
    private ushort _address;
    private ushort _value;

    public ushort Address
    {
        get => _address;
        set => SetProperty(ref _address, value);
    }

    public ushort Value
    {
        get => _value;
        set => SetProperty(ref _value, value);
    }
}
