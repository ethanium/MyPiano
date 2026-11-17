using MeltySynth;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace MyPiano;

/// <summary>
/// Software synthesizer that plays a General MIDI SoundFont (.sf2) through WASAPI.
/// Much higher quality than the default Windows "Microsoft GS Wavetable Synth".
/// </summary>
public sealed class SoundFontSynth : IMidiSink, ISampleProvider, IDisposable
{
    private const int SampleRate = 44100;

    private readonly Synthesizer _synth;
    private readonly WasapiOut _output;
    private readonly object _lock = new();
    private float[] _left = new float[1024];
    private float[] _right = new float[1024];

    /// <summary>Delay, drive, tremolo and tone, applied after the synth (not available on Windows MIDI devices).</summary>
    public EffectsChain Effects { get; } = new(SampleRate);

    public string Name { get; }
    public WaveFormat WaveFormat { get; } = WaveFormat.CreateIeeeFloatWaveFormat(SampleRate, 2);

    public SoundFontSynth(string soundFontPath)
    {
        Name = Path.GetFileName(soundFontPath);
        var settings = new SynthesizerSettings(SampleRate)
        {
            MaximumPolyphony = 128,
            EnableReverbAndChorus = true,
        };
        // .sf3 (compressed) files are decoded to PCM in memory first; .sf2 files are read directly.
        using (var stream = Sf3Loader.Open(soundFontPath))
            _synth = new Synthesizer(new SoundFont(stream), settings);
        GC.Collect();   // drop the temporary decode buffers (several times the size of the file)

        _output = new WasapiOut(AudioClientShareMode.Shared, true, 40);   // 40 ms latency
        _output.Init(this);
        _output.Play();
    }

    // Called on the audio thread.
    public int Read(float[] buffer, int offset, int count)
    {
        int frames = count / 2;
        if (_left.Length < frames)
        {
            _left = new float[frames];
            _right = new float[frames];
        }

        lock (_lock)
            _synth.Render(_left.AsSpan(0, frames), _right.AsSpan(0, frames));
        Effects.Process(_left.AsSpan(0, frames), _right.AsSpan(0, frames));

        for (int i = 0; i < frames; i++)
        {
            buffer[offset + 2 * i] = _left[i];
            buffer[offset + 2 * i + 1] = _right[i];
        }
        return count;
    }

    public void NoteOn(int channel, int note, int velocity)
    {
        lock (_lock) _synth.NoteOn(channel, note, velocity);
    }

    public void NoteOff(int channel, int note)
    {
        lock (_lock) _synth.NoteOff(channel, note);
    }

    public void ChangeInstrument(int channel, int program)
    {
        lock (_lock)
        {
            _synth.ProcessMidiMessage(channel, 0xB0, 0, 0);    // bank select MSB
            _synth.ProcessMidiMessage(channel, 0xB0, 32, 0);   // bank select LSB
            _synth.ProcessMidiMessage(channel, 0xC0, program, 0);
        }
    }

    public void Control(int channel, int controller, int value)
    {
        lock (_lock) _synth.ProcessMidiMessage(channel, 0xB0, controller, value);
    }

    public void PitchBend(int channel, int value)
    {
        value = Math.Clamp(value, 0, 16383);
        int lsb = value & 0x7F;
        int msb = (value >> 7) & 0x7F;
        lock (_lock) _synth.ProcessMidiMessage(channel, 0xE0, lsb, msb);
    }

    public void Silence()
    {
        lock (_lock) _synth.NoteOffAll(immediate: true);
    }

    public void Dispose()
    {
        _output.Stop();
        _output.Dispose();
    }
}
