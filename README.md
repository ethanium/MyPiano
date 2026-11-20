# MyPiano

A small piano, drum-pad and rhythm-box program for Windows that plays through MIDI.

- 3-octave piano played from the computer keyboard (or the mouse)
- 128 General MIDI instruments, with 4 quick-select instrument banks
- 12 percussion pads, each assignable to any of 32 drum sounds
- 3 rhythm beats (Disco, Rock, 16 Beat) with intro, two fills, an ending and a tempo slider
- Volume, transpose (key) and MIDI output device selection
- Effects (C# version): reverb, chorus, delay, drive, tremolo and tone

This repository holds **two versions of the same program**:

| Folder | What it is |
|---|---|
| [`src/MyPiano/`](src/MyPiano) | The original **Visual Basic 6** project (uses the Windows `winmm` MIDI API) |
| [`src/MyPiano.Avalonia/`](src/MyPiano.Avalonia) | A **C# / .NET 8 / Avalonia UI** port with a resizable UI and an optional built-in SoundFont synthesizer |
| [`src/tools/Sf3Tool/`](src/tools/Sf3Tool) | Command-line tool that compresses `.sf2` SoundFonts to `.sf3` |

## Controls

Both versions use the same keys.

**Piano** (the labels are drawn on the keys in the C# version)

```
 top row:     2 3 4   6 7   9 0 -        (black keys)
             Q W E R T Y U I O P [       (white keys)

 bottom row:  S D   G H J   L ;          (black keys)
             Z X C V B N M , . /         (white keys)
```

The two rows are one continuous 3-octave keyboard, not two separate ones. It starts at C on `Z`, the bottom row runs up to E on `/`, and the top row carries on from F on `Q` up to B on `[`.
In order, lowest note first: `Z S X D C V G B H N J M , L . ; /` then `Q 2 W 3 E 4 R T 6 Y 7 U I 9 O 0 P - [`.

**Drum pads**: the numpad, with Num Lock on: `/ * -`, `7 8 9`, `4 5 6`, `1 2 3`.
Click a pad to choose which drum sound it plays.

| Key | Action |
|---|---|
| `F1`-`F4` | Select instrument bank 1-4 (right-click a bank to assign the instrument you pick) |
| `F5` | Beat intro |
| `F6` / `F7` | Beat fill 1 / fill 2 |
| `F8` | Beat ending |
| `F9` | Stop the beat |

Picking a beat from the menu starts it with the intro. Fills and the ending are queued and begin when the section that is playing finishes, so they stay in time.

## C# version (Avalonia)

[`src/MyPiano.Avalonia/`](src/MyPiano.Avalonia) is built on [Avalonia](https://avaloniaui.net/).

**Requirements:** the [.NET 8 SDK](https://dotnet.microsoft.com/download).

```powershell
dotnet build src\MyPiano.Avalonia\MyPiano.Avalonia.csproj
src\MyPiano.Avalonia\bin\Debug\net8.0\MyMidi.exe
```

Audio output still uses NAudio's WASAPI backend, so sound output remains Windows-only for now; the UI itself is cross-platform.

### Sound: Windows synth or SoundFont

By default MIDI goes to a Windows MIDI device. Device 0 is normally the *Microsoft GS Wavetable Synth*, which is small and sounds dated.
For much better sound the app has a built-in synthesizer that plays a General MIDI **SoundFont**:

1. Put an `.sf2` or `.sf3` file in `src/MyPiano.Avalonia/SoundFonts/` and the app will pick it up on first run, **or**
2. Click **Device**, then **Load SoundFont (.sf2 / .sf3)…** and choose a file.

The choice is remembered in `config.json` next to the executable. Pick a normal Windows device from the same **Device** menu to switch back.

Every SoundFont in the `SoundFonts` folder is listed in the **Device** menu, so you can switch between them directly.
While a SoundFont is playing, the **Instrument** menu shows every preset it contains, not just the 128 General MIDI sounds:
variations are listed after the sound they vary, keyboard drum kits get their own submenu, and **Drum Pad Kit** picks the kit used by the drum pads and beats.

Two free General MIDI SoundFonts are included, as compressed `.sf3` files:

| File | Size | Notes |
|---|---|---|
| `FluidR3_GM.sf3` | 23 MB | Fuller sound; uses roughly 600 MB of memory when loaded |
| `GeneralUser-GS.sf3` | 8 MB | Lighter; roughly 250 MB of memory |

`.sf3` files take a second or two to decode at startup. Licenses are in [`src/MyPiano.Avalonia/SoundFonts/`](src/MyPiano.Avalonia/SoundFonts).

### Effects

Six sliders shape the sound. Their settings are saved in `config.json`.

| Slider | What it does | Works on |
|---|---|---|
| Reverb | Room/hall ambience (MIDI CC 91) | Windows MIDI devices and SoundFont synth |
| Chorus | Thicker, shimmering sound (MIDI CC 93) | Windows MIDI devices and SoundFont synth |
| Delay | Echo that bounces between left and right | SoundFont synth only |
| Drive | Soft-clip distortion | SoundFont synth only |
| Tremolo | Volume wobble (5 Hz) | SoundFont synth only |
| Tone | Left = more bass, middle = flat, right = more treble | SoundFont synth only |

Reverb and chorus are standard MIDI controls, so a Windows device only responds if its synth supports them.
The other four are built into the app's own audio output, so their sliders are greyed out while a Windows device is selected.
The code is in [`src/MyPiano.Avalonia/Effects.cs`](src/MyPiano.Avalonia/Effects.cs).

### Compressing SoundFonts (`.sf2` to `.sf3`)

GitHub rejects files over 100 MB and the original `FluidR3_GM.sf2` is about 148 MB, so it can't be pushed as is; the `.sf3` version (23 MB) can.
`src/tools/Sf3Tool` shrinks a SoundFont to about 16-26% of its size using Ogg Vorbis (lossy, quality 0.5 by default):

```powershell
dotnet run --project src\tools\Sf3Tool -c Release -- compress in.sf2 out.sf3 [quality 0.0-1.0]
dotnet run --project src\tools\Sf3Tool -c Release -- test in.sf2 out.sf3     # renders notes from both and compares
dotnet run --project src\tools\Sf3Tool -c Release -- list in.sf3                # bank, program and name of every preset
```

The tool encodes with libvorbis through [ffmpeg](https://ffmpeg.org/), which must be on your `PATH` (or set the `FFMPEG` environment variable to its location).

## Roadmap

Planned next, in order (details in [docs/ROADMAP.md](docs/ROADMAP.md)):

1. **Performance controls:** sustain pedal, pitch bend, modulation wheel, octave shift, panic button, velocity options, portamento.
2. **MIDI hardware:** MIDI input from a USB keyboard or controller, and MIDI learn for the sliders.
3. **Layering and mixing:** split and layer instruments, per-channel volume and pan, arpeggiator and chord mode.
4. **Recording:** record and play back performances, save and play `.mid` files.

## VB6 version

Open `src/MyPiano/MyPiano.vbp` in the Visual Basic 6 IDE and run it. It needs `instrument.txt` and writes its settings to `config.dat` in the same folder.
Keep the `.frx` file with `frmPiano.frm`; it holds the form's binary resources.

## Credits

- The original program came from Planet-Source-Code.com: <http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=43153&lngWId=1>.
  The author may have retained certain copyrights to this code. Please observe their request and the law by reviewing all copyright conditions at that URL.
- SoundFonts: **FluidR3 GM** by Frank Wen and Toby Smithe (MIT License) and **GeneralUser GS** by S. Christian Collins (free to use, see its license).
- C# libraries: [MeltySynth](https://github.com/sinshu/meltysynth) (SoundFont synthesizer), [NAudio](https://github.com/naudio/NAudio) (audio output), [StbVorbisSharp](https://github.com/StbSharp/StbVorbisSharp) (`.sf3` decoding), [Avalonia](https://avaloniaui.net/) (UI).
