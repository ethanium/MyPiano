# MyPiano roadmap

Planned features that copy what a real MIDI piano or controller does. Status: **planned** unless noted.
Already done: reverb, chorus, delay, drive, tremolo and tone (see the main [README](../README.md)).

"Both outputs" means it works with Windows MIDI devices and the built-in SoundFont synth.
"Synth only" means it needs the SoundFont synth.

## Phase 1: performance controls (small, both outputs)

| Feature | Notes |
|---|---|
| Sustain pedal | Hold `Space` to send CC 64. Release lets the notes die away. |
| Pitch bend | A slider that springs back to centre. Sends pitch-bend messages (0xE0). |
| Modulation wheel | A slider for vibrato (CC 1). |
| Octave shift | `Page Up` / `Page Down` and buttons. Today the Key button only transposes by semitone. |
| Panic button | Sends all-notes-off (CC 123) and all-sound-off (CC 120). Stops stuck notes without reopening the device. |
| Velocity options | Fixed, soft or hard curve, and optionally velocity by click position on the key. A computer keyboard can't sense how hard a key is pressed. |
| Portamento | Glide between notes (CC 5 and CC 65). |

## Phase 2: real MIDI hardware

| Feature | Notes |
|---|---|
| MIDI input | Play the app from a USB keyboard or controller using `winmm` `midiIn`, with the on-screen keys lighting up. Pass through the pitch wheel, mod wheel, sustain pedal and knobs. This is the biggest step up. |
| MIDI learn | Map any knob or slider on a controller to Reverb, Delay, Volume and the other sliders. |

## Phase 3: layering and mixing

| Feature | Notes |
|---|---|
| Split and layer | Two instruments at once (for example piano plus strings), or a bass on the lower keys and piano on the upper keys. |
| Per-channel mixer | Volume (CC 7) and pan (CC 10) for piano, drums and beat. |
| Arpeggiator and chord mode | One key plays a chord or an arpeggio. Reuses the beat engine's clock. |

## Phase 4: recording and playback

| Feature | Notes |
|---|---|
| Record and play back | Capture a performance, with a metronome and loop, and save it as a `.mid` file. |
| Play `.mid` files | Play standard MIDI files through the current output. |

## Order of work

1. Phase 1: sustain, pitch bend and mod wheel, octave shift, panic. About an hour.
2. Phase 2: MIDI input, then MIDI learn.
3. Phase 3 and Phase 4 as wanted.
