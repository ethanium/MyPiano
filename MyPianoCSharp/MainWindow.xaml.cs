using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Windows.Foundation;
using Windows.System;
using Windows.UI;

namespace MyPiano;

public sealed partial class MainWindow : Window
{
    // Computer-keyboard -> piano key index (36 keys, 3 octaves). Same layout as the VB6 version.
    private static readonly int[] PianoKeys =
    {
        90, 83, 88, 68, 67, 86, 71, 66, 72, 78, 74, 77, 188, 76, 190, 186, 191,
        81, 50, 87, 51, 69, 52, 82, 84, 54, 89, 55, 85, 73, 57, 79, 48, 80, 189, 219,
    };

    // Numpad keys -> drum pad index.
    private static readonly int[] DrumKeys = { 111, 106, 109, 103, 104, 105, 100, 101, 102, 97, 98, 99 };

    private static readonly string[] NoteNames = { "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B" };
    private static readonly int[] BlackSemitones = { 1, 3, 6, 8, 10 };

    private readonly MidiOut _hardware = new();
    private SoundFontSynth? _synth;
    private IMidiSink _midi;
    private readonly Settings _settings = Settings.Load();
    private readonly DispatcherTimer _beatTimer = new();

    private readonly List<string> _instruments = new();
    private readonly List<string> _drumNames = new();   // index 0 == MIDI note 35
    private IReadOnlyList<string> _devices = Array.Empty<string>();

    private readonly bool[] _held = new bool[256];
    private readonly Border[] _keyShapes = new Border[PianoKeys.Length];
    private readonly TextBlock[] _noteLabels = new TextBlock[PianoKeys.Length];
    private readonly Border[] _drumPads = new Border[12];
    private readonly int[] _whitesBefore = new int[PianoKeys.Length];
    private int _whiteCount;
    private readonly Button[] _bankButtons;

    private int _instrSelect;
    private int _baseKey;
    private int _octave;
    private int _volume = 127;

    private int _beatSelect;
    private int _beatPtr;
    private int _beatMode = -1;   // -1 stopped, else section index
    private int _beatNext = -2;   // -2 none queued

    // ---------- brushes ----------

    private static SolidColorBrush Brush(byte r, byte g, byte b) => new(Color.FromArgb(255, r, g, b));

    private static LinearGradientBrush Grad(byte r1, byte g1, byte b1, byte r2, byte g2, byte b2)
    {
        var brush = new LinearGradientBrush { StartPoint = new Point(0, 0), EndPoint = new Point(0, 1) };
        brush.GradientStops.Add(new GradientStop { Color = Color.FromArgb(255, r1, g1, b1), Offset = 0 });
        brush.GradientStops.Add(new GradientStop { Color = Color.FromArgb(255, r2, g2, b2), Offset = 1 });
        return brush;
    }

    /// <summary>Three-stop vertical gradient (top, middle, bottom) for cylindrical shading.</summary>
    private static LinearGradientBrush Grad3(byte r1, byte g1, byte b1, byte r2, byte g2, byte b2, byte r3, byte g3, byte b3)
    {
        var brush = new LinearGradientBrush { StartPoint = new Point(0, 0), EndPoint = new Point(0, 1) };
        brush.GradientStops.Add(new GradientStop { Color = Color.FromArgb(255, r1, g1, b1), Offset = 0 });
        brush.GradientStops.Add(new GradientStop { Color = Color.FromArgb(255, r2, g2, b2), Offset = 0.55 });
        brush.GradientStops.Add(new GradientStop { Color = Color.FromArgb(255, r3, g3, b3), Offset = 1 });
        return brush;
    }

    private static readonly SolidColorBrush Active = new(Colors.Gold);

    // Piano key states (cylindrical shading: bright top -> body -> darker front lip)
    private static readonly LinearGradientBrush WhiteIdle = Grad3(0xFF, 0xFF, 0xFF, 0xF6, 0xF6, 0xF9, 0xD8, 0xD8, 0xDF);
    private static readonly LinearGradientBrush BlackIdle = Grad3(0x55, 0x55, 0x5C, 0x22, 0x22, 0x26, 0x08, 0x08, 0x0A);
    private static readonly LinearGradientBrush WhiteDown = Grad3(0xFF, 0xE6, 0xA0, 0xFF, 0xC2, 0x4D, 0xF0, 0x9A, 0x00);
    private static readonly LinearGradientBrush BlackDown = Grad3(0xFF, 0xD0, 0x50, 0xFF, 0xB0, 0x20, 0xE0, 0x7A, 0x00);
    private static readonly SolidColorBrush KeyBorderWhite = Brush(0xA8, 0xA8, 0xB2);
    private static readonly SolidColorBrush KeyBorderBlack = Brush(0x00, 0x00, 0x00);
    private static readonly SolidColorBrush WhiteLip = Brush(0xC4, 0xC4, 0xCC);
    private static readonly SolidColorBrush BlackLip = Brush(0x00, 0x00, 0x00);

    // Drum pad states
    private static readonly SolidColorBrush PadIdleBg = Brush(0x24, 0x24, 0x2A);
    private static readonly SolidColorBrush PadIdleFg = Brush(0xC8, 0xC8, 0xD0);
    private static readonly SolidColorBrush PadDownFg = Brush(0x14, 0x14, 0x14);

    // Beveled hardware-button gradients
    private static readonly LinearGradientBrush BtnNormal = Grad3(0x34, 0x34, 0x3B, 0x2A, 0x2A, 0x30, 0x20, 0x20, 0x25);
    private static readonly LinearGradientBrush BtnHover = Grad3(0x40, 0x40, 0x48, 0x34, 0x34, 0x3B, 0x28, 0x28, 0x2E);
    private static readonly LinearGradientBrush BtnPressed = Grad3(0x1C, 0x1C, 0x21, 0x22, 0x22, 0x27, 0x28, 0x28, 0x2D);
    private static readonly SolidColorBrush BtnBorder = Brush(0x3C, 0x3C, 0x44);

    public MainWindow()
    {
        InitializeComponent();
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1120, 720));

        _bankButtons = new[] { BtnBank0, BtnBank1, BtnBank2, BtnBank3 };

        LoadInstrumentNames();
        BuildPiano();
        BuildDrumPads();
        WireControls();

        // Devices
        _devices = MidiOut.GetDeviceNames();
        if (_settings.DeviceId >= _devices.Count) _settings.DeviceId = 0;
        _midi = _hardware;
        if (string.IsNullOrEmpty(_settings.SoundFontPath) || !File.Exists(_settings.SoundFontPath))
        {
            // First run (or file moved): use a SoundFont found in SoundFonts\ next to the exe.
            // Prefer .sf2 (loads fastest) over .sf3 (compressed, decoded at startup).
            var dir = Path.Combine(AppContext.BaseDirectory, "SoundFonts");
            var found = Directory.Exists(dir)
                ? Directory.EnumerateFiles(dir)
                    .Where(f => f.EndsWith(".sf2", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".sf3", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(f => f.EndsWith(".sf3", StringComparison.OrdinalIgnoreCase)).ThenBy(f => f)
                    .FirstOrDefault()
                : null;
            if (found is not null)
            {
                _settings.SoundFontPath = found;
                _settings.UseSoundFont = true;
            }
        }
        if (_settings.UseSoundFont && File.Exists(_settings.SoundFontPath) && TryUseSoundFont(_settings.SoundFontPath!))
        {
            // using the SoundFont synth
        }
        else
        {
            _settings.UseSoundFont = false;
            OpenDevice(_settings.DeviceId);
        }

        SelectBeat(0, start: false);
        SelectInstrument(_settings.Banks[0]);

        Root.AddHandler(UIElement.KeyDownEvent, new KeyEventHandler(OnKeyDown), true);
        Root.AddHandler(UIElement.KeyUpEvent, new KeyEventHandler(OnKeyUp), true);
        Root.Loaded += (_, _) =>
        {
            Root.Focus(FocusState.Programmatic);
            // First layout pass: SizeChanged may have fired before arrange completed.
            LayoutPiano(PianoHost.ActualWidth, PianoHost.ActualHeight);
        };

        Activated += (_, e) =>
        {
            if (e.WindowActivationState == WindowActivationState.Deactivated) ReleaseAll();
        };
        Closed += (_, _) =>
        {
            _beatTimer.Stop();
            _synth?.Dispose();
            _hardware.Close();
            _settings.Save();
        };
    }

    // ---------- setup ----------

    private void LoadInstrumentNames()
    {
        try
        {
            var lines = File.ReadAllLines(Path.Combine(AppContext.BaseDirectory, "Instrument.txt"));
            for (int i = 0; i < 128 && i < lines.Length; i++)
                _instruments.Add(lines[i].Split('\t')[1]);
            // line 128 is blank; drums follow (notes 35..66)
            for (int i = 129; i < lines.Length; i++)
            {
                var parts = lines[i].Split('\t');
                if (parts.Length > 1) _drumNames.Add(parts[1]);
            }
        }
        catch (Exception ex)
        {
            ShowError("Could not load Instrument.txt: " + ex.Message);
        }
        while (_instruments.Count < 128) _instruments.Add($"Instrument {_instruments.Count + 1}");
        while (_drumNames.Count < 32) _drumNames.Add($"Drum {_drumNames.Count + 35}");
    }

    private static string KeyLabel(int vk) => vk switch
    {
        188 => ",",
        190 => ".",
        186 => ";",
        191 => "/",
        189 => "-",
        219 => "[",
        111 => "/",
        106 => "*",
        109 => "-",
        >= 96 and <= 105 => (vk - 96).ToString(),   // numpad 0-9
        _ => ((char)vk).ToString(),                 // letters and top-row digits
    };

    private void BuildPiano()
    {
        int whiteCount = 0;
        var blacks = new List<int>();

        for (int i = 0; i < PianoKeys.Length; i++)
        {
            bool black = BlackSemitones.Contains(i % 12);
            int index = i;
            _whitesBefore[i] = whiteCount;

            // Front lip: a thin darker strip along the bottom edge to fake depth.
            var lip = new Border
            {
                VerticalAlignment = VerticalAlignment.Bottom,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Height = 8,
                Background = black ? BlackLip : WhiteLip,
                CornerRadius = new CornerRadius(0, 0, black ? 4 : 5, black ? 4 : 5),
                IsHitTestVisible = false,
            };

            var keyLabel = new TextBlock
            {
                Text = KeyLabel(PianoKeys[i]),
                FontSize = 14,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                Foreground = new SolidColorBrush(black ? Colors.White : Colors.Black),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 10),
                IsHitTestVisible = false,
            };

            // Note-name label (e.g. C4, F#3) at the top of the key; updated on key/octave change.
            var noteLabel = new TextBlock
            {
                Text = NoteNameFor(i),
                FontSize = 11,
                FontWeight = Microsoft.UI.Text.FontWeights.Normal,
                Foreground = new SolidColorBrush(black ? Color.FromArgb(255, 0xB0, 0xB0, 0xB8) : Color.FromArgb(255, 0x70, 0x70, 0x78)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 6, 0, 0),
                IsHitTestVisible = false,
            };
            _noteLabels[i] = noteLabel;

            // Host grid so the lip and labels can stack inside the key border.
            var content = new Grid();
            content.Children.Add(noteLabel);
            content.Children.Add(keyLabel);
            content.Children.Add(lip);

            var shape = new Border
            {
                Background = black ? BlackIdle : WhiteIdle,
                BorderBrush = black ? KeyBorderBlack : KeyBorderWhite,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(0, 0, black ? 5 : 6, black ? 5 : 6),
                Child = content,
            };
            shape.PointerPressed += (s, e) => { ((UIElement)s).CapturePointer(e.Pointer); SetKey(index, true); };
            shape.PointerReleased += (s, e) => { ((UIElement)s).ReleasePointerCapture(e.Pointer); SetKey(index, false); };
            shape.PointerCaptureLost += (_, _) => SetKey(index, false);
            _keyShapes[i] = shape;

            if (black)
            {
                blacks.Add(i);
            }
            else
            {
                PianoCanvas.Children.Add(shape);
                whiteCount++;
            }
        }
        foreach (int i in blacks)
            PianoCanvas.Children.Add(_keyShapes[i]);   // added last => drawn on top

        _whiteCount = whiteCount;
        PianoHost.SizeChanged += (_, e) => LayoutPiano(e.NewSize.Width, e.NewSize.Height);
    }

    /// <summary>Scales the keyboard to fill the available space, keeping realistic key proportions.</summary>
    private void LayoutPiano(double width, double height)
    {
        if (_whiteCount == 0 || width <= 0 || height <= 0) return;

        double whiteW = width / _whiteCount;
        // Real piano proportions: white key 23.5 x 150 mm (1:6.4), black key 13.7 x 90 mm.
        // Fill available height but clamp so keys never look absurdly long on tall windows.
        double whiteH = Math.Clamp(height, 90, whiteW * 6.4);
        double blackW = whiteW * 0.58;   // 13.7 / 23.5
        double blackH = whiteH * 0.60;   // 90 / 150
        double lipH = Math.Clamp(whiteH * 0.05, 4, 9);
        double font = Math.Clamp(whiteW * 0.34, 8, 18);
        double noteFont = Math.Clamp(whiteW * 0.22, 7, 13);

        for (int i = 0; i < _keyShapes.Length; i++)
        {
            bool black = BlackSemitones.Contains(i % 12);
            var shape = _keyShapes[i];
            shape.Width = black ? blackW : whiteW;
            shape.Height = black ? blackH : whiteH;
            Canvas.SetLeft(shape, black ? _whitesBefore[i] * whiteW - blackW / 2 : _whitesBefore[i] * whiteW);

            if (_noteLabels[i] is { } noteLabel) noteLabel.FontSize = noteFont;

            if (shape.Child is Grid content)
            {
                foreach (var child in content.Children)
                {
                    if (child is TextBlock label && !ReferenceEquals(label, _noteLabels[i])) label.FontSize = font;
                    else if (child is Border lip)
                    {
                        lip.Height = lipH;
                        lip.CornerRadius = new CornerRadius(0, 0, black ? 4 : 5, black ? 4 : 5);
                    }
                }
            }
        }
        PianoCanvas.Width = width;
        PianoCanvas.Height = whiteH;
    }

    private void BuildDrumPads()
    {
        for (int i = 0; i < 12; i++)
        {
            int index = i;
            var pad = new Border
            {
                // Initial square; recomputed on size changes in LayoutDrumPads.
                Width = 64,
                Height = 64,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                CornerRadius = new CornerRadius(9),
                Background = PadIdleBg,
                BorderBrush = Brush(0x3A, 0x3A, 0x46),
                BorderThickness = new Thickness(1),
                Child = new TextBlock
                {
                    Text = KeyLabel(DrumKeys[i]),
                    FontSize = 16,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    Foreground = PadIdleFg,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false,
                },
            };
            Grid.SetRow(pad, i / 3);
            Grid.SetColumn(pad, i % 3);

            pad.PointerPressed += (s, e) => { ((UIElement)s).CapturePointer(e.Pointer); SetDrum(index, true); };
            pad.PointerReleased += (s, e) => { ((UIElement)s).ReleasePointerCapture(e.Pointer); SetDrum(index, false); };
            pad.PointerCaptureLost += (_, _) => SetDrum(index, false);
            pad.RightTapped += (s, e) => { e.Handled = true; ShowDrumMenu(pad, index); };

            _drumPads[i] = pad;
            DrumGrid.Children.Add(pad);
            UpdatePadTooltip(i);
        }

        // Keep pads square as the panel resizes.
        DrumPanel.SizeChanged += (_, e) => LayoutDrumPads(e.NewSize.Width, e.NewSize.Height);
    }

    /// <summary>Sizes the drum pads as squares that fit the panel (3 columns x 4 rows).</summary>
    private void LayoutDrumPads(double panelWidth, double panelHeight)
    {
        const double gap = 6;          // RowSpacing/ColumnSpacing
        const double padBorder = 28;   // DrumPanel Padding (14) * 2
        const double labelSpace = 30;  // section label + spacing above the grid

        double availW = panelWidth - padBorder - gap * 2;      // 3 columns, 2 gaps
        double availH = panelHeight - padBorder - labelSpace - gap * 3; // 4 rows, 3 gaps
        if (availW <= 0 || availH <= 0) return;

        double byWidth = availW / 3;
        double byHeight = availH / 4;
        double size = Math.Clamp(Math.Min(byWidth, byHeight), 36, 120);

        foreach (var pad in _drumPads)
        {
            pad.Width = size;
            pad.Height = size;
            if (pad.Child is TextBlock t) t.FontSize = Math.Clamp(size * 0.3, 11, 22);
        }
    }

    private void WireControls()
    {
        BtnInstrument.Click += (_, _) =>
            BuildInstrumentFlyout(i => SelectInstrument(i)).ShowAt(BtnInstrument);

        for (int b = 0; b < 4; b++)
        {
            int bank = b;
            var btn = _bankButtons[b];
            btn.Click += (_, _) => SelectInstrument(_settings.Banks[bank]);
            btn.RightTapped += (_, e) =>
            {
                e.Handled = true;
                BuildInstrumentFlyout(i =>
                {
                    _settings.Banks[bank] = i;
                    SelectInstrument(i);
                }).ShowAt(btn);
            };
        }

        BtnBeat.Click += (_, _) =>
        {
            var flyout = new MenuFlyout();
            for (int i = 0; i < Beat.All.Length; i++)
            {
                int index = i;
                var item = new MenuFlyoutItem { Text = Beat.All[i].Name };
                item.Click += (_, _) => SelectBeat(index, start: true);
                flyout.Items.Add(item);
            }
            flyout.ShowAt(BtnBeat);
        };

        foreach (var btn in new[] { BtnIntro, BtnFill1, BtnFill2, BtnEnd, BtnStop })
        {
            var b = btn;
            b.Click += (_, _) => BeatControl(int.Parse((string)b.Tag));
        }

        SldTempo.ValueChanged += (_, e) => { TxtTempo.Text = ((int)e.NewValue).ToString(); UpdateBeatInterval(); };
        SldVolume.ValueChanged += (_, e) => { _volume = (int)e.NewValue; TxtVolume.Text = _volume.ToString(); };

        SldReverb.Value = _settings.Reverb;
        SldChorus.Value = _settings.Chorus;
        TxtReverb.Text = _settings.Reverb.ToString();
        TxtChorus.Text = _settings.Chorus.ToString();
        SldReverb.ValueChanged += (_, e) => { _settings.Reverb = (int)e.NewValue; TxtReverb.Text = ((int)e.NewValue).ToString(); ApplyEffects(); };
        SldChorus.ValueChanged += (_, e) => { _settings.Chorus = (int)e.NewValue; TxtChorus.Text = ((int)e.NewValue).ToString(); ApplyEffects(); };

        SldDelay.Value = _settings.Delay;
        SldDrive.Value = _settings.Drive;
        SldTremolo.Value = _settings.Tremolo;
        SldTone.Value = _settings.Tone;
        TxtDelay.Text = _settings.Delay.ToString();
        TxtDrive.Text = _settings.Drive.ToString();
        TxtTremolo.Text = _settings.Tremolo.ToString();
        TxtTone.Text = _settings.Tone.ToString();
        SldDelay.ValueChanged += (_, e) => { _settings.Delay = (int)e.NewValue; TxtDelay.Text = ((int)e.NewValue).ToString(); ApplyEffects(); };
        SldDrive.ValueChanged += (_, e) => { _settings.Drive = (int)e.NewValue; TxtDrive.Text = ((int)e.NewValue).ToString(); ApplyEffects(); };
        SldTremolo.ValueChanged += (_, e) => { _settings.Tremolo = (int)e.NewValue; TxtTremolo.Text = ((int)e.NewValue).ToString(); ApplyEffects(); };
        SldTone.ValueChanged += (_, e) => { _settings.Tone = (int)e.NewValue; TxtTone.Text = ((int)e.NewValue).ToString(); ApplyEffects(); };

        BtnOctDown.Click += (_, _) => ShiftOctave(-1);
        BtnOctUp.Click += (_, _) => ShiftOctave(1);

        BtnKey.Click += (_, _) =>
        {
            var flyout = new MenuFlyout();
            for (int i = 0; i < 12; i++)
            {
                int index = i;
                var item = new MenuFlyoutItem { Text = NoteNames[i] };
                item.Click += (_, _) =>
                {
                    _baseKey = index;
                    BtnKey.Content = NoteNames[index];
                    ReleaseAll();
                    _midi.Silence();   // the original reopened the device to cut off hanging notes
                    RefreshNoteLabels();
                    ApplyEffects();    // reopening a hardware device resets its controllers
                };
                flyout.Items.Add(item);
            }
            flyout.ShowAt(BtnKey);
        };

        BtnDevice.Click += (_, _) =>
        {
            var flyout = new MenuFlyout();
            for (int i = 0; i < _devices.Count; i++)
            {
                int index = i;
                var item = new MenuFlyoutItem { Text = _devices[i] };
                item.Click += (_, _) =>
                {
                    OpenDevice(index);
                    SelectInstrument(_instrSelect);
                };
                flyout.Items.Add(item);
            }

            flyout.Items.Add(new MenuFlyoutSeparator());
            if (!string.IsNullOrEmpty(_settings.SoundFontPath) && File.Exists(_settings.SoundFontPath))
            {
                var current = new MenuFlyoutItem { Text = "SoundFont synth: " + Path.GetFileName(_settings.SoundFontPath) };
                current.Click += (_, _) =>
                {
                    if (TryUseSoundFont(_settings.SoundFontPath!)) SelectInstrument(_instrSelect);
                };
                flyout.Items.Add(current);
            }
            var browse = new MenuFlyoutItem { Text = "Load SoundFont (.sf2 / .sf3)…" };
            browse.Click += async (_, _) => await PickSoundFontAsync();
            flyout.Items.Add(browse);

            flyout.ShowAt(BtnDevice);
        };

        _beatTimer.Tick += (_, _) => BeatTick();

        // Beveled hardware look for every button.
        foreach (var b in new Button[]
                 {
                     BtnDevice, BtnInstrument, BtnBank0, BtnBank1, BtnBank2, BtnBank3,
                     BtnBeat, BtnIntro, BtnFill1, BtnFill2, BtnEnd, BtnStop,
                     BtnKey, BtnOctDown, BtnOctUp,
                 })
        {
            StyleBevel(b);
        }
    }

    /// <summary>Applies a beveled gradient + hover/press feedback to a button.</summary>
    private static void StyleBevel(Button b)
    {
        b.Background = BtnNormal;
        b.BorderBrush = BtnBorder;
        b.BorderThickness = new Thickness(1);
        b.CornerRadius = new CornerRadius(6);
        b.Foreground = Brush(0xE6, 0xE6, 0xEA);
        b.Padding = new Thickness(10, 6, 10, 6);

        b.PointerEntered += (_, _) => { if (b.IsEnabled) b.Background = BtnHover; };
        b.PointerExited += (_, _) => { if (b.IsEnabled) b.Background = BtnNormal; };
        b.PointerPressed += (_, _) => { if (b.IsEnabled) b.Background = BtnPressed; };
        b.PointerReleased += (_, _) => { if (b.IsEnabled) b.Background = BtnHover; };
    }

    // ---------- menus ----------

    private MenuFlyout BuildInstrumentFlyout(Action<int> onPick)
    {
        var flyout = new MenuFlyout();
        for (int family = 0; family < 16; family++)
        {
            var sub = new MenuFlyoutSubItem { Text = $"{family * 8 + 1}-{family * 8 + 8}: {_instruments[family * 8]}…" };
            for (int j = 0; j < 8; j++)
            {
                int index = family * 8 + j;
                var item = new MenuFlyoutItem { Text = _instruments[index] };
                item.Click += (_, _) => onPick(index);
                sub.Items.Add(item);
            }
            flyout.Items.Add(sub);
        }
        return flyout;
    }

    private void ShowDrumMenu(FrameworkElement anchor, int pad)
    {
        var flyout = new MenuFlyout();
        for (int i = 0; i < _drumNames.Count; i++)
        {
            int note = i + 35;
            var item = new MenuFlyoutItem { Text = _drumNames[i] };
            item.Click += (_, _) =>
            {
                _settings.Drums[pad] = note;
                UpdatePadTooltip(pad);
            };
            flyout.Items.Add(item);
        }
        flyout.ShowAt(anchor);
    }

    private void UpdatePadTooltip(int pad)
    {
        int idx = _settings.Drums[pad] - 35;
        string name = idx >= 0 && idx < _drumNames.Count ? _drumNames[idx] : $"Note {_settings.Drums[pad]}";
        ToolTipService.SetToolTip(_drumPads[pad], name);
    }

    // ---------- MIDI actions ----------

    private void OpenDevice(int id)
    {
        try
        {
            ReleaseAll();
            _synth?.Dispose();
            _synth = null;
            _settings.UseSoundFont = false;
            _settings.DeviceId = id;
            _hardware.Open(id);
            _midi = _hardware;
            BtnDevice.Content = id < _devices.Count ? _devices[id] : $"Device {id}";
        }
        catch (Exception ex)
        {
            BtnDevice.Content = "(none)";
            ShowError(ex.Message);
        }
    }

    /// <summary>Switches output to the built-in SoundFont synth. Returns false (and shows why) on failure.</summary>
    private bool TryUseSoundFont(string path)
    {
        try
        {
            ReleaseAll();
            var synth = new SoundFontSynth(path);   // load first so a bad file leaves the old output intact
            _synth?.Dispose();
            _synth = synth;
            _hardware.Close();
            _midi = synth;
            _settings.SoundFontPath = path;
            _settings.UseSoundFont = true;
            BtnDevice.Content = "SoundFont: " + synth.Name;
            return true;
        }
        catch (Exception ex)
        {
            ShowError($"Could not load SoundFont: {ex.Message}");
            return false;
        }
    }

    private async Task PickSoundFontAsync()
    {
        var picker = new Windows.Storage.Pickers.FileOpenPicker();
        WinRT.Interop.InitializeWithWindow.Initialize(picker, WinRT.Interop.WindowNative.GetWindowHandle(this));
        picker.FileTypeFilter.Add(".sf2");
        picker.FileTypeFilter.Add(".sf3");
        var file = await picker.PickSingleFileAsync();
        if (file is null) return;
        if (TryUseSoundFont(file.Path)) SelectInstrument(_instrSelect);
    }

    private void SelectInstrument(int index)
    {
        _instrSelect = index;
        _midi.ChangeInstrument(0, index);
        ApplyEffects();   // also covers a freshly opened device or synth
        BtnInstrument.Content = _instruments[index];
    }

    /// <summary>Sends the reverb (CC 91) and chorus (CC 93) levels to the piano and drum channels.</summary>
    private void ApplyEffects()
    {
        foreach (int ch in new[] { 0, 9 })
        {
            _midi.Control(ch, 91, _settings.Reverb);
            _midi.Control(ch, 93, _settings.Chorus);
        }

        // Delay, drive, tremolo and tone only exist on the built-in synth.
        bool synth = _midi is SoundFontSynth;
        foreach (var s in new[] { SldDelay, SldDrive, SldTremolo, SldTone }) s.IsEnabled = synth;
        if (_midi is SoundFontSynth sf)
        {
            sf.Effects.Delay = _settings.Delay;
            sf.Effects.Drive = _settings.Drive;
            sf.Effects.Tremolo = _settings.Tremolo;
            sf.Effects.Tone = _settings.Tone;
        }
    }

    private int NoteFor(int i) => i + _baseKey + 48 + _octave * 12;

    /// <summary>Scientific pitch name (e.g. C4, F#3) for key index i.</summary>
    private string NoteNameFor(int i)
    {
        int note = NoteFor(i);
        return NoteNames[note % 12] + (note / 12 - 1);
    }

    /// <summary>Refreshes every key's note-name label after a key or octave change.</summary>
    private void RefreshNoteLabels()
    {
        for (int i = 0; i < _noteLabels.Length; i++)
            if (_noteLabels[i] is not null) _noteLabels[i].Text = NoteNameFor(i);
    }

    private void SetKey(int i, bool down)
    {
        bool black = BlackSemitones.Contains(i % 12);
        var shape = _keyShapes[i];
        bool wasDown = shape.Tag is true;
        if (down == wasDown) return;
        shape.Tag = down;

        if (down)
        {
            shape.Background = black ? BlackDown : WhiteDown;
            _midi.NoteOn(0, NoteFor(i), _volume);
        }
        else
        {
            shape.Background = black ? BlackIdle : WhiteIdle;
            _midi.NoteOff(0, NoteFor(i));
        }
    }

    private void SetDrum(int i, bool down)
    {
        var pad = _drumPads[i];
        if (down)
        {
            pad.Background = Active;
            if (pad.Child is TextBlock t) t.Foreground = PadDownFg;
            _midi.NoteOn(9, _settings.Drums[i], 127);
        }
        else
        {
            pad.Background = PadIdleBg;
            if (pad.Child is TextBlock t) t.Foreground = PadIdleFg;
            _midi.NoteOff(9, _settings.Drums[i]);
        }
    }

    private void ShiftOctave(int delta)
    {
        ReleaseAll();
        _midi.Silence();
        _octave = Math.Clamp(_octave + delta, -2, 2);
        TxtOct.Text = _octave > 0 ? "+" + _octave : _octave.ToString();
        RefreshNoteLabels();
        ApplyEffects();
    }

    private void ReleaseAll()
    {
        for (int k = 0; k < _held.Length; k++)
        {
            if (!_held[k]) continue;
            _held[k] = false;
            int p = Array.IndexOf(PianoKeys, k);
            if (p >= 0) SetKey(p, false);
            int d = Array.IndexOf(DrumKeys, k);
            if (d >= 0) SetDrum(d, false);
        }
    }

    // ---------- keyboard ----------

    private void OnKeyDown(object sender, KeyRoutedEventArgs e)
    {
        int k = (int)e.Key;
        if (k < 0 || k > 255 || _held[k]) return;   // ignore auto-repeat

        bool used = false;
        if (k >= (int)VirtualKey.F1 && k <= (int)VirtualKey.F4)
        {
            int bank = k - (int)VirtualKey.F1;
            SelectInstrument(_settings.Banks[bank]);
            _bankButtons[bank].Background = Active;
            used = true;
        }
        else if (k >= (int)VirtualKey.F5 && k <= (int)VirtualKey.F9)
        {
            BeatControl(k - (int)VirtualKey.F5);
            used = true;
        }

        int p = Array.IndexOf(PianoKeys, k);
        if (p >= 0) { SetKey(p, true); used = true; }
        int d = Array.IndexOf(DrumKeys, k);
        if (d >= 0) { SetDrum(d, true); used = true; }

        if (used) { _held[k] = true; e.Handled = true; }
    }

    private void OnKeyUp(object sender, KeyRoutedEventArgs e)
    {
        int k = (int)e.Key;
        if (k < 0 || k > 255) return;
        _held[k] = false;

        if (k >= (int)VirtualKey.F1 && k <= (int)VirtualKey.F4)
            _bankButtons[k - (int)VirtualKey.F1].Background = BtnNormal;

        int p = Array.IndexOf(PianoKeys, k);
        if (p >= 0) SetKey(p, false);
        int d = Array.IndexOf(DrumKeys, k);
        if (d >= 0) SetDrum(d, false);
    }

    // ---------- rhythm ----------

    private void SelectBeat(int index, bool start)
    {
        _beatSelect = index;
        _beatPtr = 0;
        _beatMode = start ? 0 : -1;   // picking a beat from the menu starts it with the intro, like the original
        _beatNext = -2;
        BtnBeat.Content = Beat.All[index].Name;
        SldTempo.Value = 0;
        UpdateBeatInterval();
        if (start) _beatTimer.Start();
    }

    private void UpdateBeatInterval()
    {
        _beatTimer.Interval = TimeSpan.FromMilliseconds(Beat.All[_beatSelect].Rate + SldTempo.Value);
    }

    /// <summary>0 intro, 1 fill 1, 2 fill 2, 3 ending, 4 stop.</summary>
    private void BeatControl(int index)
    {
        _beatNext = index switch { 0 => 0, 1 => 2, 2 => 3, 3 => 4, _ => -1 };
        if (index < 4 && _beatMode < 0)
        {
            _beatMode = _beatNext;
            _beatNext = -2;
            _beatPtr = 0;
            _beatTimer.Start();
        }
    }

    private void BeatTick()
    {
        if (_beatMode < 0) { _beatTimer.Stop(); return; }

        var section = Beat.All[_beatSelect].Sections[_beatMode];
        if (_beatPtr >= section.Length) _beatPtr = 0;

        foreach (int note in section[_beatPtr])
            _midi.NoteOn(9, note, _volume);

        _beatPtr++;
        if (_beatPtr >= section.Length)
        {
            if (_beatNext < -1)
            {
                _beatMode = _beatMode switch { 0 or 2 or 3 => 1, 4 => -1, _ => _beatMode };
            }
            else
            {
                _beatMode = _beatNext;
                _beatNext = -2;
            }
            _beatPtr = 0;
        }
    }

    // ---------- misc ----------

    private async void ShowError(string message)
    {
        if (Content?.XamlRoot is null) return;
        await new ContentDialog
        {
            Title = "My MIDI",
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = Content.XamlRoot,
        }.ShowAsync();
    }
}