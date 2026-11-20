namespace MyPiano;

/// <summary>
/// Stereo effects chain applied to the SoundFont synth's output: Drive, Tone, Tremolo, then Delay.
/// All amounts are 0-127 (Tone is a tilt: 64 = flat). An effect at its neutral setting is skipped entirely.
/// Amounts may be changed from the UI thread while <see cref="Process"/> runs on the audio thread.
/// </summary>
public sealed class EffectsChain
{
    private readonly int _sampleRate;

    // Written by the UI thread, read by the audio thread (single float writes are atomic).
    private volatile float _delay, _drive, _tremolo, _tone;

    public int Delay { set => _delay = Norm(value); }
    public int Drive { set => _drive = Norm(value); }
    public int Tremolo { set => _tremolo = Norm(value); }
    /// <summary>0 = dark (bass), 64 = flat, 127 = bright (treble).</summary>
    public int Tone { set => _tone = (Math.Clamp(value, 0, 127) - 64) / 64f; }

    private static float Norm(int v) => Math.Clamp(v, 0, 127) / 127f;

    // Tone: one-pole low-pass splits bass from treble.
    private float _lowL, _lowR;
    private readonly float _toneCoeff;

    // Tremolo: 5 Hz LFO.
    private double _phase;
    private readonly double _phaseStep;

    // Delay: ping-pong echo, 375 ms.
    private readonly float[] _dlyL, _dlyR;
    private int _pos;
    private const float Feedback = 0.45f;

    public EffectsChain(int sampleRate)
    {
        _sampleRate = sampleRate;
        _toneCoeff = 1f - MathF.Exp(-2f * MathF.PI * 800f / sampleRate);
        _phaseStep = 2 * Math.PI * 5.0 / sampleRate;
        int len = (int)(0.375 * sampleRate);
        _dlyL = new float[len];
        _dlyR = new float[len];
    }

    public void Process(Span<float> left, Span<float> right)
    {
        float drive = _drive, tone = _tone, trem = _tremolo, delay = _delay;
        bool doDrive = drive > 0.005f, doTone = MathF.Abs(tone) > 0.01f, doTrem = trem > 0.005f;
        bool doDelay = delay > 0.005f;
        if (!(doDrive || doTone || doTrem || doDelay)) return;

        float gain = 1f + drive * 12f;
        float norm = 1f / MathF.Tanh(gain);
        float lowGain = 1f - 0.8f * tone, highGain = 1f + 0.8f * tone;

        for (int i = 0; i < left.Length; i++)
        {
            float l = left[i], r = right[i];

            if (doDrive)
            {
                // Soft-clipping; blended with the dry signal so low settings are subtle.
                l += (MathF.Tanh(l * gain) * norm - l) * drive;
                r += (MathF.Tanh(r * gain) * norm - r) * drive;
            }

            if (doTone)
            {
                _lowL += _toneCoeff * (l - _lowL);
                _lowR += _toneCoeff * (r - _lowR);
                l = _lowL * lowGain + (l - _lowL) * highGain;
                r = _lowR * lowGain + (r - _lowR) * highGain;
            }

            if (doTrem)
            {
                float g = 1f - trem * (0.5f + 0.5f * (float)Math.Sin(_phase));
                l *= g;
                r *= g;
                _phase += _phaseStep;
                if (_phase > 2 * Math.PI) _phase -= 2 * Math.PI;
            }

            if (doDelay)
            {
                float outL = _dlyL[_pos], outR = _dlyR[_pos];
                _dlyL[_pos] = (l + r) * 0.5f + outR * Feedback;   // ping-pong: each side feeds the other
                _dlyR[_pos] = outL * Feedback;
                if (++_pos >= _dlyL.Length) _pos = 0;
                l += outL * delay;
                r += outR * delay;
            }

            left[i] = Math.Clamp(l, -1f, 1f);
            right[i] = Math.Clamp(r, -1f, 1f);
        }
    }
}
