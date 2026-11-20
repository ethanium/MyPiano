using System.Diagnostics;
using System.Globalization;
using MeltySynth;
using MyPiano;

// sf3tool compress <in.sf2> <out.sf3> [quality 0.0-1.0, default 0.5]
// sf3tool test     <original.sf2> <compressed.sf3>
// sf3tool list     <in.sf2|in.sf3>
//
// Compression encodes every sample with the real libvorbis encoder (through ffmpeg, which must be on
// PATH or named by the FFMPEG environment variable). libvorbis streams are what other SF3 tools
// (MuseScore, FluidSynth) expect, and they come back sample-aligned.
if (args is ["list", var listPath])
    return List(listPath);

if (args.Length < 3 || args[0] is not ("compress" or "test"))
{
    Console.WriteLine("""
        Usage:
          sf3tool compress <in.sf2> <out.sf3> [quality 0.0-1.0]   (default 0.5)
          sf3tool test     <original.sf2> <compressed.sf3>
          sf3tool list     <in.sf2|in.sf3>                         (bank, program, name of every preset)
        """);
    return 1;
}

return args[0] == "compress"
    ? Compress(args[1], args[2], args.Length > 3 ? float.Parse(args[3], CultureInfo.InvariantCulture) : 0.5f)
    : Test(args[1], args[2]);

// Prints every preset as "bank<TAB>program<TAB>name"; programs are 1-based like the app's menus,
// bank 128 holds the drum kits.
static int List(string path)
{
    SoundFont sf;
    using (var stream = Sf3Loader.Open(path))
        sf = new SoundFont(stream);
    foreach (var p in sf.Presets.OrderBy(p => p.BankNumber).ThenBy(p => p.PatchNumber))
        Console.WriteLine($"{p.BankNumber}\t{p.PatchNumber + 1}\t{p.Name.Trim()}");
    return 0;
}

static int Compress(string input, string output, float quality)
{
    var watch = Stopwatch.StartNew();
    var sf = SfFile.Read(File.ReadAllBytes(input));
    var headers = sf.Headers;
    int count = headers.Length - 1;   // last record is the "EOS" terminator
    if (headers.Take(count).Any(h => h.IsCompressed))
    {
        Console.Error.WriteLine("Input is already compressed.");
        return 1;
    }

    // Encode each distinct sample once, in parallel.
    var seg = sf.SmplSegment;
    var unique = new Dictionary<(uint, uint, uint), int>();
    var jobs = new List<(uint Start, uint End, uint Rate)>();
    var slot = new int[count];
    for (int i = 0; i < count; i++)
    {
        var key = (headers[i].Start, headers[i].End, headers[i].SampleRate);
        if (!unique.TryGetValue(key, out int j)) { j = jobs.Count; unique[key] = j; jobs.Add(key); }
        slot[i] = j;
    }

    var encoded = new byte[jobs.Count][];
    var length = new uint[jobs.Count];   // decoded length in samples (>= original: libvorbis pads the tail)
    int done = 0;
    Parallel.For(0, jobs.Count, j =>
    {
        var (start, end, rate) = jobs[j];
        var pcm = new short[end - start];
        Buffer.BlockCopy(seg.Array!, seg.Offset + (int)start * 2, pcm, 0, pcm.Length * 2);
        try
        {
            encoded[j] = Encode(pcm, (int)rate, quality);
            length[j] = (uint)Sf3Loader.Decode(encoded[j]).Length;
            Verify(pcm, encoded[j]);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Sample #{j} ({pcm.Length} samples @ {rate} Hz): {ex.Message}", ex);
        }
        int n = Interlocked.Increment(ref done);
        if (n % 100 == 0) Console.Write($"\r  encoded {n}/{jobs.Count} samples");
    });
    Console.WriteLine($"\r  encoded {jobs.Count}/{jobs.Count} samples      ");

    long total = encoded.Sum(e => (long)e.Length);
    var smpl = new byte[total];
    var offsets = new uint[jobs.Count];
    long cursor = 0;
    for (int j = 0; j < encoded.Length; j++)
    {
        offsets[j] = (uint)cursor;
        encoded[j].CopyTo(smpl, cursor);
        cursor += encoded[j].Length;
    }

    var outHeaders = (SfSampleHeader[])headers.Clone();
    for (int i = 0; i < count; i++)
    {
        var h = headers[i];
        int j = slot[i];
        outHeaders[i] = h with
        {
            Start = offsets[j],
            End = offsets[j] + (uint)encoded[j].Length,
            // SF3 stores loop points relative to the start of the sample.
            StartLoop = h.StartLoop >= h.Start ? h.StartLoop - h.Start : 0,
            EndLoop = h.EndLoop >= h.Start ? h.EndLoop - h.Start : 0,
            Type = (ushort)(h.Type | SfSampleHeader.OggFlag),
        };
    }

    using (var fs = File.Create(output))
        sf.Write(fs, smpl, SfFile.SerializeHeaders(outHeaders), 3, 0);

    long before = new FileInfo(input).Length, after = new FileInfo(output).Length;
    Console.WriteLine($"{Path.GetFileName(input)}: {before / 1048576.0:F1} MB -> {after / 1048576.0:F1} MB " +
                      $"({100.0 * after / before:F0}%) in {watch.Elapsed.TotalSeconds:F1}s, quality {quality}");
    return 0;
}

// Encodes 16-bit mono PCM to an Ogg Vorbis stream with libvorbis (via ffmpeg, piped in and out).
static byte[] Encode(short[] pcm, int rate, float quality)
{
    var ffmpeg = Environment.GetEnvironmentVariable("FFMPEG") ?? "ffmpeg";
    var q = (quality * 10).ToString("0.##", CultureInfo.InvariantCulture);   // ffmpeg's -q:a is on a 0-10 scale
    var psi = new ProcessStartInfo(ffmpeg,
        $"-v error -f s16le -ar {rate} -ac 1 -i - -c:a libvorbis -q:a {q} -f ogg -")
    {
        RedirectStandardInput = true,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true,
    };

    Process process;
    try { process = Process.Start(psi)!; }
    catch (Exception ex) { throw new InvalidOperationException($"could not start '{ffmpeg}' ({ex.Message}); install ffmpeg or set FFMPEG", ex); }

    using (process)
    {
        var stdout = new MemoryStream();
        var copyOut = process.StandardOutput.BaseStream.CopyToAsync(stdout);
        var readErr = process.StandardError.ReadToEndAsync();

        var raw = new byte[pcm.Length * 2];
        Buffer.BlockCopy(pcm, 0, raw, 0, raw.Length);
        using (var stdin = process.StandardInput.BaseStream)
            stdin.Write(raw, 0, raw.Length);

        copyOut.Wait();
        process.WaitForExit();
        if (process.ExitCode != 0 || stdout.Length == 0)
            throw new InvalidOperationException($"ffmpeg failed: {readErr.Result.Trim()}");
        return stdout.ToArray();
    }
}

// Decodes with the same decoder the app uses and checks the result is the original (plus at most a
// little silence libvorbis pads onto the tail) and starts at the same place.
static void Verify(short[] original, byte[] ogg)
{
    var decoded = Sf3Loader.Decode(ogg);
    if (decoded.Length < original.Length || decoded.Length > original.Length + 4096)
        throw new InvalidOperationException($"decoded {decoded.Length} samples, expected about {original.Length}");

    double signal = 0, noise = 0;
    for (int i = 0; i < original.Length; i++)
    {
        double a = original[i], d = a - decoded[i];
        signal += a * a;
        noise += d * d;
    }
    // A shifted stream would have a very poor SNR here; even noisy samples stay well above this.
    if (signal >= 1 && 10 * Math.Log10(signal / Math.Max(noise, 1)) < 4)
        throw new InvalidOperationException("decoded audio does not line up with the original");
}

// Renders the same notes through the original and the compressed SoundFont and compares them.
static int Test(string originalPath, string compressedPath)
{
    var watch = Stopwatch.StartNew();
    using var originalStream = File.OpenRead(originalPath);
    var original = new SoundFont(originalStream);
    Console.WriteLine($"original loaded in {watch.Elapsed.TotalSeconds:F1}s");

    watch.Restart();
    using var decoded = Sf3Loader.Open(compressedPath);
    var compressed = new SoundFont(decoded);
    Console.WriteLine($"compressed loaded (decoded + parsed) in {watch.Elapsed.TotalSeconds:F1}s, " +
                      $"{GC.GetTotalMemory(false) / 1048576} MB managed");

    // (program, channel, key, label)
    var cases = new[]
    {
        (0, 0, 60, "Grand Piano C4"), (0, 0, 84, "Grand Piano C6"), (24, 0, 55, "Nylon Guitar"),
        (40, 0, 69, "Violin A4"), (56, 0, 64, "Trumpet"), (19, 0, 48, "Church Organ (looped)"),
        (42, 0, 40, "Cello (low)"), (0, 9, 38, "Drums: snare"), (0, 9, 42, "Drums: closed hi-hat"),
    };

    int bad = 0;
    foreach (var (program, channel, key, label) in cases)
    {
        var a = Render(original, program, channel, key);
        var b = Render(compressed, program, channel, key);
        double peakA = a.Max(Math.Abs), peakB = b.Max(Math.Abs);
        if (peakA < 0.001)
        {
            // This SoundFont plays nothing for that note/program (silent in the original too).
            Console.WriteLine($"  skip {label,-24} original is silent");
            continue;
        }
        double snr0 = Snr(a, b, 0);

        // Best alignment within +/-2000 samples, to tell a time shift from real distortion.
        int bestLag = 0;
        double best = snr0;
        for (int lag = -2000; lag <= 2000; lag++)
        {
            double s = Snr(a, b, lag);
            if (s > best) { best = s; bestLag = lag; }
        }
        // Alignment (lag 0) is the strict check. Noise-like sounds such as hi-hats score a low waveform
        // SNR with any lossy codec while still sounding the same, so the SNR bar is modest.
        bool ok = peakB > 0.001 && snr0 > 8 && bestLag == 0;
        if (!ok) bad++;
        Console.WriteLine($"  {(ok ? "ok  " : "FAIL")} {label,-24} peak {peakA:F3} vs {peakB:F3}   " +
                          $"SNR {snr0,5:F1} dB   best {best,5:F1} dB at lag {bestLag}");
    }
    Console.WriteLine(bad == 0 ? "All sounds match." : $"{bad} sound(s) differ too much.");
    return bad == 0 ? 0 : 1;
}

// SNR of b against a, with b shifted by lag samples (b[i + lag] vs a[i]).
static double Snr(float[] a, float[] b, int lag)
{
    double signal = 0, noise = 0;
    for (int i = 0; i < a.Length; i++)
    {
        int j = i + lag;
        if (j < 0 || j >= b.Length) continue;
        signal += a[i] * a[i];
        noise += (a[i] - b[j]) * (a[i] - b[j]);
    }
    return 10 * Math.Log10(Math.Max(signal, 1e-12) / Math.Max(noise, 1e-12));
}

static float[] Render(SoundFont font, int program, int channel, int key)
{
    var synth = new Synthesizer(font, new SynthesizerSettings(44100));
    if (channel != 9) synth.ProcessMidiMessage(channel, 0xC0, program, 0);
    synth.NoteOn(channel, key, 100);
    var left = new float[44100];
    var right = new float[44100];
    synth.Render(left, right);
    return left;
}
