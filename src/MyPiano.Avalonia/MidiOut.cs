using System.Runtime.InteropServices;

namespace MyPiano;

/// <summary>Anything that can play MIDI notes: a hardware/OS device or the built-in SoundFont synth.</summary>
public interface IMidiSink
{
    void NoteOn(int channel, int note, int velocity);
    void NoteOff(int channel, int note);
    /// <summary>Selects a program (0-127) from a bank (bank select MSB, 0 = General MIDI).</summary>
    void ChangeInstrument(int channel, int bank, int program);
    /// <summary>Sends a MIDI control change (e.g. 91 = reverb send, 93 = chorus send).</summary>
    void Control(int channel, int controller, int value);
    /// <summary>Sends a pitch bend message (0..16383, 8192 is center).</summary>
    void PitchBend(int channel, int value);
    /// <summary>Stops any hanging notes.</summary>
    void Silence();
}

/// <summary>Thin wrapper over the Windows winmm MIDI output API (port of Vbmidi.bas).</summary>
public sealed class MidiOut : IMidiSink, IDisposable
{
    private int _deviceId;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct MIDIOUTCAPS
    {
        public ushort wMid;
        public ushort wPid;
        public uint vDriverVersion;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szPname;
        public ushort wTechnology;
        public ushort wVoices;
        public ushort wNotes;
        public ushort wChannelMask;
        public uint dwSupport;
    }

    [DllImport("winmm.dll")]
    private static extern uint midiOutGetNumDevs();

    [DllImport("winmm.dll", CharSet = CharSet.Unicode)]
    private static extern uint midiOutGetDevCapsW(UIntPtr deviceId, ref MIDIOUTCAPS caps, uint size);

    [DllImport("winmm.dll")]
    private static extern uint midiOutOpen(out IntPtr handle, uint deviceId, IntPtr callback, IntPtr instance, uint flags);

    [DllImport("winmm.dll")]
    private static extern uint midiOutClose(IntPtr handle);

    [DllImport("winmm.dll")]
    private static extern uint midiOutShortMsg(IntPtr handle, uint msg);

    private IntPtr _handle;

    public bool IsOpen => _handle != IntPtr.Zero;

    public static IReadOnlyList<string> GetDeviceNames()
    {
        var names = new List<string>();
        uint count = midiOutGetNumDevs();
        for (uint i = 0; i < count; i++)
        {
            var caps = new MIDIOUTCAPS();
            names.Add(midiOutGetDevCapsW((UIntPtr)i, ref caps, (uint)Marshal.SizeOf<MIDIOUTCAPS>()) == 0
                ? caps.szPname
                : $"Device {i}");
        }
        return names;
    }

    public void Open(int deviceId)
    {
        Close();
        _deviceId = deviceId;
        if (midiOutOpen(out _handle, (uint)deviceId, IntPtr.Zero, IntPtr.Zero, 0) != 0)
        {
            _handle = IntPtr.Zero;
            throw new InvalidOperationException($"Open MIDI Out failed (device {deviceId}).");
        }
    }

    public void Close()
    {
        if (_handle == IntPtr.Zero) return;
        midiOutClose(_handle);
        _handle = IntPtr.Zero;
    }

    private void Send(int status, int channel, int data1, int data2 = 0)
    {
        if (_handle == IntPtr.Zero) return;
        uint msg = (uint)(status | channel | ((data1 & 0x7F) << 8) | ((data2 & 0x7F) << 16));
        midiOutShortMsg(_handle, msg);
    }

    public void NoteOn(int channel, int note, int velocity) => Send(0x90, channel, note, velocity);

    public void NoteOff(int channel, int note) => Send(0x80, channel, note);

    public void PitchBend(int channel, int value)
    {
        value = Math.Clamp(value, 0, 16383);
        int lsb = value & 0x7F;
        int msb = (value >> 7) & 0x7F;
        Send(0xE0, channel, lsb, msb);
    }

    public void ChangeInstrument(int channel, int bank, int program)
    {
        Send(0xB0, channel, 0, bank);   // bank select MSB
        Send(0xB0, channel, 32, 0);  // bank select LSB
        Send(0xC0, channel, program);
    }

    public void Control(int channel, int controller, int value) => Send(0xB0, channel, controller, value);

    // Closing and reopening the device cuts off any sounding notes (as the VB6 version did).
    public void Silence()
    {
        if (IsOpen) Open(_deviceId);
    }

    public void Dispose() => Close();
}
