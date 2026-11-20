using StbVorbisSharp;

namespace MyPiano;

/// <summary>
/// Opens a SoundFont for MeltySynth. A .sf3 (Ogg Vorbis-compressed samples) is decoded in memory and
/// rewritten as a plain .sf2 stream, because MeltySynth only understands uncompressed samples.
/// </summary>
public static class Sf3Loader
{
    private const int Padding = 46;   // SF2 requires 46 zero samples after each sample

    public static Stream Open(string path)
    {
        if (!path.EndsWith(".sf3", StringComparison.OrdinalIgnoreCase))
            return File.OpenRead(path);

        var sf = SfFile.Read(File.ReadAllBytes(path));
        var headers = sf.Headers;
        int count = headers.Length - 1;   // last record is the "EOS" terminator

        // Decode each distinct sample once (some headers can share the same range).
        var unique = new Dictionary<(uint, uint, bool), int>();
        var jobs = new List<(uint Start, uint End, bool Ogg)>();
        var slot = new int[count];
        for (int i = 0; i < count; i++)
        {
            var key = (headers[i].Start, headers[i].End, headers[i].IsCompressed);
            if (!unique.TryGetValue(key, out int j))
            {
                j = jobs.Count;
                unique[key] = j;
                jobs.Add(key);
            }
            slot[i] = j;
        }

        var pcm = new short[jobs.Count][];
        var seg = sf.SmplSegment;   // spans can't be captured by the parallel lambda
        Parallel.For(0, jobs.Count, j =>
        {
            var (start, end, ogg) = jobs[j];
            pcm[j] = ogg
                ? Decode(seg.AsSpan((int)start, (int)(end - start)))
                : ReadPcm(seg, start, end);
        });

        // Lay the decoded samples out one after another and fix up the headers.
        long total = 0;
        foreach (var p in pcm) total += p.Length + Padding;
        var offsets = new uint[jobs.Count];
        var lengths = new uint[jobs.Count];
        long cursor = 0;
        for (int j = 0; j < pcm.Length; j++)
        {
            offsets[j] = (uint)cursor;
            lengths[j] = (uint)pcm[j].Length;
            cursor += pcm[j].Length + Padding;
        }

        var fixedHeaders = (SfSampleHeader[])headers.Clone();
        for (int i = 0; i < count; i++)
        {
            var h = headers[i];
            int j = slot[i];
            uint start = offsets[j];
            uint len = lengths[j];
            // Compressed samples store loop points relative to the sample start; SF2 wants absolute.
            uint loopStart = h.IsCompressed ? start + h.StartLoop : h.StartLoop - h.Start + start;
            uint loopEnd = h.IsCompressed ? start + h.EndLoop : h.EndLoop - h.Start + start;
            fixedHeaders[i] = h with
            {
                Start = start,
                End = start + len,
                StartLoop = Math.Min(loopStart, start + len),
                EndLoop = Math.Min(loopEnd, start + len),
                Type = (ushort)(h.Type & ~SfSampleHeader.OggFlag),
            };
        }

        var padding = new byte[Padding * 2];
        var ms = new MemoryStream((int)(total * 2 + 4 * 1024 * 1024));
        sf.Write(ms, total * 2, stream =>
        {
            for (int j = 0; j < pcm.Length; j++)
            {
                stream.Write(System.Runtime.InteropServices.MemoryMarshal.AsBytes(pcm[j].AsSpan()));
                stream.Write(padding);
                pcm[j] = null!;   // free each decoded sample as soon as it has been written
            }
        }, SfFile.SerializeHeaders(fixedHeaders), 2, 1);
        ms.Position = 0;
        return ms;
    }

    private static short[] ReadPcm(ArraySegment<byte> smpl, uint startSample, uint endSample)
    {
        var pcm = new short[endSample - startSample];
        Buffer.BlockCopy(smpl.Array!, smpl.Offset + (int)startSample * 2, pcm, 0, pcm.Length * 2);
        return pcm;
    }

    /// <summary>Decodes one Ogg Vorbis stream to 16-bit mono PCM.</summary>
    public static short[] Decode(ReadOnlySpan<byte> ogg)
    {
        // StbVorbisSharp keeps a raw pointer into the buffer it is given while decoding, so the buffer must
        // not move: allocate it on the pinned object heap (an ordinary array can be relocated by the GC,
        // which showed up as access violations on large SoundFonts).
        var data = GC.AllocateUninitializedArray<byte>(ogg.Length, pinned: true);
        ogg.CopyTo(data);

        // It also crashes when called from several threads at once, so decoding is serialized. It is fast
        // enough that this isn't a bottleneck.
        lock (DecodeLock)
        {
            using var vorbis = Vorbis.FromMemory(data);
            int channels = vorbis.Channels;
            var result = new List<short>((int)(vorbis.LengthInSeconds * vorbis.SampleRate) + 4096);   // avoid regrowth
            while (true)
            {
                vorbis.SubmitBuffer();
                int frames = vorbis.Decoded;   // frames (samples per channel) now in SongBuffer
                if (frames <= 0) break;
                var buf = vorbis.SongBuffer;
                for (int i = 0; i < frames; i++)   // SoundFont samples are mono; keep channel 0
                    result.Add(buf[i * channels]);
            }
            return result.ToArray();
        }
    }

    private static readonly object DecodeLock = new();
}
