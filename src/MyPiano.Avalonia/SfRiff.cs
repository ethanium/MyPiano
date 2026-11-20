using System.Buffers.Binary;
using System.Text;

namespace MyPiano;

/// <summary>One 46-byte SoundFont sample header (the "shdr" chunk).</summary>
public struct SfSampleHeader
{
    public byte[] Name;          // 20 bytes
    public uint Start, End, StartLoop, EndLoop, SampleRate;
    public byte Pitch;
    public sbyte Correction;
    public ushort Link, Type;

    /// <summary>SF3 flag: the sample is Ogg Vorbis data and Start/End are byte offsets.</summary>
    public const ushort OggFlag = 0x10;
    public const int Size = 46;

    public readonly bool IsCompressed => (Type & OggFlag) != 0;
}

/// <summary>
/// Minimal SoundFont (RIFF "sfbk") reader/writer: just enough to swap the sample data ("smpl")
/// and the sample headers ("shdr") while copying every other chunk untouched.
/// </summary>
public sealed class SfFile
{
    private byte[] _raw = Array.Empty<byte>();
    private int _smplOffset, _smplLength;

    public List<(string Id, byte[] Data)> Info { get; } = new();
    public List<(string Id, byte[] Data)> Pdta { get; } = new();

    public ReadOnlySpan<byte> Smpl => _raw.AsSpan(_smplOffset, _smplLength);

    /// <summary>The sample data as a segment of the file buffer (no copy), for use across threads.</summary>
    public ArraySegment<byte> SmplSegment => new(_raw, _smplOffset, _smplLength);

    public SfSampleHeader[] Headers
    {
        get
        {
            var shdr = Pdta.FirstOrDefault(c => c.Id == "shdr").Data
                ?? throw new InvalidDataException("SoundFont has no sample headers (shdr).");
            var result = new SfSampleHeader[shdr.Length / SfSampleHeader.Size];
            for (int i = 0; i < result.Length; i++)
            {
                var s = shdr.AsSpan(i * SfSampleHeader.Size, SfSampleHeader.Size);
                result[i] = new SfSampleHeader
                {
                    Name = s[..20].ToArray(),
                    Start = BinaryPrimitives.ReadUInt32LittleEndian(s[20..]),
                    End = BinaryPrimitives.ReadUInt32LittleEndian(s[24..]),
                    StartLoop = BinaryPrimitives.ReadUInt32LittleEndian(s[28..]),
                    EndLoop = BinaryPrimitives.ReadUInt32LittleEndian(s[32..]),
                    SampleRate = BinaryPrimitives.ReadUInt32LittleEndian(s[36..]),
                    Pitch = s[40],
                    Correction = (sbyte)s[41],
                    Link = BinaryPrimitives.ReadUInt16LittleEndian(s[42..]),
                    Type = BinaryPrimitives.ReadUInt16LittleEndian(s[44..]),
                };
            }
            return result;
        }
    }

    public static byte[] SerializeHeaders(SfSampleHeader[] headers)
    {
        var buf = new byte[headers.Length * SfSampleHeader.Size];
        for (int i = 0; i < headers.Length; i++)
        {
            var h = headers[i];
            var s = buf.AsSpan(i * SfSampleHeader.Size, SfSampleHeader.Size);
            h.Name.AsSpan(0, 20).CopyTo(s);
            BinaryPrimitives.WriteUInt32LittleEndian(s[20..], h.Start);
            BinaryPrimitives.WriteUInt32LittleEndian(s[24..], h.End);
            BinaryPrimitives.WriteUInt32LittleEndian(s[28..], h.StartLoop);
            BinaryPrimitives.WriteUInt32LittleEndian(s[32..], h.EndLoop);
            BinaryPrimitives.WriteUInt32LittleEndian(s[36..], h.SampleRate);
            s[40] = h.Pitch;
            s[41] = (byte)h.Correction;
            BinaryPrimitives.WriteUInt16LittleEndian(s[42..], h.Link);
            BinaryPrimitives.WriteUInt16LittleEndian(s[44..], h.Type);
        }
        return buf;
    }

    private static string Id(byte[] b, int pos) => Encoding.ASCII.GetString(b, pos, 4);

    public static SfFile Read(byte[] raw)
    {
        if (raw.Length < 12 || Id(raw, 0) != "RIFF" || Id(raw, 8) != "sfbk")
            throw new InvalidDataException("Not a SoundFont file.");

        var f = new SfFile { _raw = raw };
        int end = (int)Math.Min(raw.Length, 8L + BinaryPrimitives.ReadUInt32LittleEndian(raw.AsSpan(4)));
        int pos = 12;
        while (pos + 8 <= end)
        {
            string id = Id(raw, pos);
            int size = (int)BinaryPrimitives.ReadUInt32LittleEndian(raw.AsSpan(pos + 4));
            int body = pos + 8;
            if (id == "LIST")
            {
                string type = Id(raw, body);
                int p = body + 4, e = Math.Min(raw.Length, body + size);
                while (p + 8 <= e)
                {
                    string cid = Id(raw, p);
                    int csize = (int)BinaryPrimitives.ReadUInt32LittleEndian(raw.AsSpan(p + 4));
                    int cbody = p + 8;
                    switch (type)
                    {
                        case "INFO": f.Info.Add((cid, raw.AsSpan(cbody, csize).ToArray())); break;
                        case "pdta": f.Pdta.Add((cid, raw.AsSpan(cbody, csize).ToArray())); break;
                        case "sdta" when cid == "smpl": f._smplOffset = cbody; f._smplLength = csize; break;
                    }
                    p = cbody + csize + (csize & 1);
                }
            }
            pos = body + size + (size & 1);
        }
        if (f._smplLength == 0) throw new InvalidDataException("SoundFont has no sample data (smpl).");
        return f;
    }

    private static long ChunkTotal(long len) => 8 + len + (len & 1);

    private static void WriteChunk(Stream s, string id, ReadOnlySpan<byte> data)
    {
        Span<byte> head = stackalloc byte[8];
        Encoding.ASCII.GetBytes(id, head);
        BinaryPrimitives.WriteUInt32LittleEndian(head[4..], (uint)data.Length);
        s.Write(head);
        s.Write(data);
        if ((data.Length & 1) != 0) s.WriteByte(0);
    }

    private static void WriteListHeader(Stream s, string type, long size)
    {
        Span<byte> head = stackalloc byte[12];
        Encoding.ASCII.GetBytes("LIST", head);
        BinaryPrimitives.WriteUInt32LittleEndian(head[4..], (uint)size);
        Encoding.ASCII.GetBytes(type, head[8..]);
        s.Write(head);
    }

    /// <summary>Writes this SoundFont with <paramref name="smpl"/> and <paramref name="shdr"/> replacing the originals.</summary>
    public void Write(Stream s, ReadOnlySpan<byte> smpl, byte[] shdr, ushort versionMajor, ushort versionMinor)
    {
        var copy = smpl.ToArray();   // the callback form can't capture a span
        Write(s, copy.Length, stream => stream.Write(copy), shdr, versionMajor, versionMinor);
    }

    /// <summary>
    /// Like <see cref="Write(Stream, ReadOnlySpan{byte}, byte[], ushort, ushort)"/>, but the sample data is
    /// produced by <paramref name="writeSmpl"/> (exactly <paramref name="smplLength"/> bytes), so a large
    /// block never has to exist in memory as one piece.
    /// </summary>
    public void Write(Stream s, long smplLength, Action<Stream> writeSmpl, byte[] shdr, ushort versionMajor, ushort versionMinor)
    {
        var info = Info.Select(c => c.Id == "ifil" && c.Data.Length >= 4
            ? (c.Id, Data: Version(c.Data, versionMajor, versionMinor))
            : c).ToList();
        var pdta = Pdta.Select(c => c.Id == "shdr" ? (c.Id, Data: shdr) : c).ToList();

        long infoSize = 4 + info.Sum(c => ChunkTotal(c.Data.Length));
        long sdtaSize = 4 + ChunkTotal(smplLength);
        long pdtaSize = 4 + pdta.Sum(c => ChunkTotal(c.Data.Length));
        long riffSize = 4 + (8 + infoSize) + (8 + sdtaSize) + (8 + pdtaSize);
        if (riffSize > uint.MaxValue) throw new InvalidDataException("SoundFont is larger than 4 GB.");

        Span<byte> riff = stackalloc byte[12];
        Encoding.ASCII.GetBytes("RIFF", riff);
        BinaryPrimitives.WriteUInt32LittleEndian(riff[4..], (uint)riffSize);
        Encoding.ASCII.GetBytes("sfbk", riff[8..]);
        s.Write(riff);

        WriteListHeader(s, "INFO", infoSize);
        foreach (var c in info) WriteChunk(s, c.Id, c.Data);
        WriteListHeader(s, "sdta", sdtaSize);
        Span<byte> smplHead = stackalloc byte[8];
        Encoding.ASCII.GetBytes("smpl", smplHead);
        BinaryPrimitives.WriteUInt32LittleEndian(smplHead[4..], (uint)smplLength);
        s.Write(smplHead);
        writeSmpl(s);
        if ((smplLength & 1) != 0) s.WriteByte(0);
        WriteListHeader(s, "pdta", pdtaSize);
        foreach (var c in pdta) WriteChunk(s, c.Id, c.Data);
    }

    private static byte[] Version(byte[] original, ushort major, ushort minor)
    {
        var copy = (byte[])original.Clone();
        BinaryPrimitives.WriteUInt16LittleEndian(copy.AsSpan(0), major);
        BinaryPrimitives.WriteUInt16LittleEndian(copy.AsSpan(2), minor);
        return copy;
    }
}
