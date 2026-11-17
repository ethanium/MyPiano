using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;

namespace MyPiano;

public partial class MainWindow : Window
{
    // Computer-keyboard -> piano key index (36 keys, 3 octaves).
    private static readonly Key[] PianoKeys =
    {
        Key.Z, Key.S, Key.X, Key.D, Key.C, Key.V, Key.G, Key.B, Key.H, Key.N, Key.J, Key.M,
        Key.OemComma, Key.L, Key.OemPeriod, Key.OemSemicolon, Key.OemQuestion,
        Key.Q, Key.D2, Key.W, Key.D3, Key.E, Key.D4, Key.R, Key.T, Key.D6, Key.Y, Key.D7, Key.U,
        Key.I, Key.D9, Key.O, Key.D0, Key.P, Key.OemMinus, Key.OemOpenBrackets,
    };

    // Numpad keys -> drum pad index.
    private static readonly Key[] DrumKeys =
    {
        Key.Divide, Key.Multiply, Key.Subtract, Key.NumPad7, Key.NumPad8, Key.NumPad9,
        Key.NumPad4, Key.NumPad5, Key.NumPad6, Key.NumPad1, Key.NumPad2, Key.NumPad3,
    };

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

    private readonly HashSet<Key> _held = new();
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

    private static SolidColorBrush Brush(byte r, byte g, byte b) => new(Color.FromRgb(r, g, b));

    /// <summary>Three-stop vertical gradient (top, middle, bottom) for cylindrical shading.</summary>
    private static LinearGradientBrush Grad3(byte r1, byte g1, byte b1, byte r2, byte g2, byte b2, byte r3, byte g3, byte b3)
    {
        var brush = new LinearGradientBrush
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
        };
        brush.GradientStops.Add(new GradientStop(Color.FromRgb(r1, g1, b1), 0));
        brush.GradientStops.Add(new GradientStop(Color.FromRgb(r2, g2, b2), 0.55));
        brush.GradientStops.Add(new GradientStop(Color.FromRgb(r3, g3, b3), 1));
        return brush;
    }

    private static readonly IBrush Active = Brushes.Gold;

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
    private static readonly SolidColorBrush BtnFg = Brush(0xE6, 0xE6, 0xEA);

    public MainWindow()
    {
        InitializeComponent();

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

        Root.AddHandler(KeyDownEvent, new EventHandler<KeyEventArgs>(OnKeyDown),
            RoutingStrategies.Tunnel | RoutingStrategies.Bubble, true);
        Root.AddHandler(KeyUpEvent, new EventHandler<KeyEventArgs>(OnKeyUp),
            RoutingStrategies.Tunnel | RoutingStrategies.Bubble, true);

        Loaded += (_, _) =>
        {
            Root.Focus();
            LayoutPiano(PianoHost.Bounds.Width, PianoHost.Bounds.Height);
            LayoutDrumPads();
        };

        // A stuck key would keep sounding after the window loses focus.
        Deactivated += (_, _) => ReleaseAll();
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

    private static string KeyLabel(Key k) => k switch
    {
        Key.OemComma => ",",
        Key.OemPeriod => ".",
        Key.OemSemicolon => ";",
        Key.OemQuestion => "/",
        Key.OemMinus => "-",
        Key.OemOpenBrackets => "[",
        Key.Divide => "/",
        Key.Multiply => "*",
        Key.Subtract => "-",
        >= Key.D0 and <= Key.D9 => (k - Key.D0).ToString(),
        >= Key.NumPad0 and <= Key.NumPad9 => (k - Key.NumPad0).ToString(),
        _ => k.ToString(),
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
                FontWeight = FontWeight.SemiBold,
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
                FontWeight = FontWeight.Normal,
                Foreground = new SolidColorBrush(black ? Color.FromRgb(0xB0, 0xB0, 0xB8) : Color.FromRgb(0x70, 0x70, 0x78)),
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
            shape.PointerPressed += (s, e) =>
            {
                e.Pointer.Capture((InputElement)s!);
                SetKey(index, true);
            };
            shape.PointerReleased += (s, e) =>
            {
                e.Pointer.Capture(null);
                SetKey(index, false);
            };
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
                    FontWeight = FontWeight.SemiBold,
                    Foreground = PadIdleFg,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false,
                },
            };
            Grid.SetRow(pad, i / 3);
            Grid.SetColumn(pad, i % 3);

            pad.PointerPressed += (s, e) =>
            {
                if (e.GetCurrentPoint(pad).Properties.IsRightButtonPressed)
                {
                    ShowDrumMenu(pad, index);
                    return;
                }
                e.Pointer.Capture((InputElement)s!);
                SetDrum(index, true);
            };
            pad.PointerReleased += (_, _) => SetDrum(index, false);
            pad.PointerCaptureLost += (_, _) => SetDrum(index, false);

            _drumPads[i] = pad;
            DrumGrid.Children.Add(pad);
            UpdatePadTooltip(i);
        }

        // Keep pads square as the panel resizes.
        DrumPanel.SizeChanged += (_, _) => LayoutDrumPads();
    }

    /// <summary>Sizes the drum pads as squares that fit the panel (3 columns x 4 rows).</summary>
    private void LayoutDrumPads()
    {
        const double gap = 6;          // grid gaps
        const double padBorder = 28;   // DrumPanel Padding (14) * 2
        const double labelSpace = 30;  // section label + spacing above the grid

        double panelWidth = DrumPanel.Bounds.Width;
        double panelHeight = DrumPanel.Bounds.Height;
        if (panelWidth <= 0 || panelHeight <= 0) return;

        double availW = panelWidth - padBorder - gap * 2;                // 3 columns, 2 gaps
        double availH = panelHeight - padBorder - labelSpace - gap * 3;  // 4 rows, 3 gaps
        if (availW <= 0 || availH <= 0) return;

        double size = Math.Clamp(Math.Min(availW / 3, availH / 4), 36, 120);

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
            BuildInstrumentMenu(i => SelectInstrument(i)).ShowAt(BtnInstrument);

        for (int b = 0; b < 4; b++)
        {
            int bank = b;
            var btn = _bankButtons[b];
            btn.Click += (_, _) => SelectInstrument(_settings.Banks[bank]);
            btn.PointerPressed += (s, e) =>
            {
                if (!e.GetCurrentPoint(btn).Properties.IsRightButtonPressed) return;
                BuildInstrumentMenu(i =>
                {
                    _settings.Banks[bank] = i;
                    SelectInstrument(i);
                }).ShowAt(btn);
            };
        }

        BtnBeat.Click += (_, _) =>
        {
            var menu = new MenuFlyout();
            for (int i = 0; i < Beat.All.Length; i++)
            {
                int index = i;
                var item = new MenuItem { Header = Beat.All[index].Name };
                item.Click += (_, _) => SelectBeat(index, start: true);
                menu.Items.Add(item);
            }
            menu.ShowAt(BtnBeat);
        };

        foreach (var btn in new[] { BtnIntro, BtnFill1, BtnFill2, BtnEnd, BtnStop })
        {
            var b = btn;
            b.Click += (_, _) => BeatControl(int.Parse((string)b.Tag!));
        }

        SldTempo.ValueChanged += (_, _) => { TxtTempo.Text = ((int)SldTempo.Value).ToString(); UpdateBeatInterval(); };
        SldVolume.ValueChanged += (_, _) => { _volume = (int)SldVolume.Value; TxtVolume.Text = _volume.ToString(); };

        SldReverb.Value = _settings.Reverb;
        SldChorus.Value = _settings.Chorus;
        TxtReverb.Text = _settings.Reverb.ToString();
        TxtChorus.Text = _settings.Chorus.ToString();
        SldReverb.ValueChanged += (_, _) => { _settings.Reverb = (int)SldReverb.Value; TxtReverb.Text = ((int)SldReverb.Value).ToString(); ApplyEffects(); };
        SldChorus.ValueChanged += (_, _) => { _settings.Chorus = (int)SldChorus.Value; TxtChorus.Text = ((int)SldChorus.Value).ToString(); ApplyEffects(); };

        SldDelay.Value = _settings.Delay;
        SldDrive.Value = _settings.Drive;
        SldTremolo.Value = _settings.Tremolo;
        SldTone.Value = _settings.Tone;
        TxtDelay.Text = _settings.Delay.ToString();
        TxtDrive.Text = _settings.Drive.ToString();
        TxtTremolo.Text = _settings.Tremolo.ToString();
        TxtTone.Text = _settings.Tone.ToString();
        SldDelay.ValueChanged += (_, _) => { _settings.Delay = (int)SldDelay.Value; TxtDelay.Text = ((int)SldDelay.Value).ToString(); ApplyEffects(); };
        SldDrive.ValueChanged += (_, _) => { _settings.Drive = (int)SldDrive.Value; TxtDrive.Text = ((int)SldDrive.Value).ToString(); ApplyEffects(); };
        SldTremolo.ValueChanged += (_, _) => { _settings.Tremolo = (int)SldTremolo.Value; TxtTremolo.Text = ((int)SldTremolo.Value).ToString(); ApplyEffects(); };
        SldTone.ValueChanged += (_, _) => { _settings.Tone = (int)SldTone.Value; TxtTone.Text = ((int)SldTone.Value).ToString(); ApplyEffects(); };

        BtnOctDown.Click += (_, _) => ShiftOctave(-1);
        BtnOctUp.Click += (_, _) => ShiftOctave(1);

        BtnKey.Click += (_, _) =>
        {
            var menu = new MenuFlyout();
            for (int i = 0; i < 12; i++)
            {
                int index = i;
                var item = new MenuItem { Header = NoteNames[index] };
                item.Click += (_, _) =>
                {
                    _baseKey = index;
                    BtnKey.Content = NoteNames[index];
                    ReleaseAll();
                    _midi.Silence();   // the original reopened the device to cut off hanging notes
                    RefreshNoteLabels();
                    ApplyEffects();    // reopening a hardware device resets its controllers
                };
                menu.Items.Add(item);
            }
            menu.ShowAt(BtnKey);
        };

        BtnDevice.Click += (_, _) =>
        {
            var menu = new MenuFlyout();
            for (int i = 0; i < _devices.Count; i++)
            {
                int index = i;
                var item = new MenuItem { Header = _devices[i] };
                item.Click += (_, _) =>
                {
                    OpenDevice(index);
                    SelectInstrument(_instrSelect);
                };
                menu.Items.Add(item);
            }

            menu.Items.Add(new Separator());
            if (!string.IsNullOrEmpty(_settings.SoundFontPath) && File.Exists(_settings.SoundFontPath))
            {
                var current = new MenuItem { Header = "SoundFont synth: " + Path.GetFileName(_settings.SoundFontPath) };
                current.Click += (_, _) =>
                {
                    if (TryUseSoundFont(_settings.SoundFontPath!)) SelectInstrument(_instrSelect);
                };
                menu.Items.Add(current);
            }
            var browse = new MenuItem { Header = "Load SoundFont (.sf2 / .sf3)…" };
            browse.Click += async (_, _) => await PickSoundFontAsync();
            menu.Items.Add(browse);

            menu.ShowAt(BtnDevice);
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
        b.Foreground = BtnFg;
        b.Padding = new Thickness(10, 6, 10, 6);

        b.PointerEntered += (_, _) => { if (b.IsEnabled) b.Background = BtnHover; };
        b.PointerExited += (_, _) => { if (b.IsEnabled) b.Background = BtnNormal; };
        b.PointerPressed += (_, _) => { if (b.IsEnabled) b.Background = BtnPressed; };
        b.PointerReleased += (_, _) => { if (b.IsEnabled) b.Background = BtnHover; };
    }

    // ---------- menus ----------

    private MenuFlyout BuildInstrumentMenu(Action<int> onPick)
    {
        var menu = new MenuFlyout();
        for (int family = 0; family < 16; family++)
        {
            var sub = new MenuItem { Header = $"{family * 8 + 1}-{family * 8 + 8}: {_instruments[family * 8]}…" };
            for (int j = 0; j < 8; j++)
            {
                int index = family * 8 + j;
                var item = new MenuItem { Header = _instruments[index] };
                item.Click += (_, _) => onPick(index);
                sub.Items.Add(item);
            }
            menu.Items.Add(sub);
        }
        return menu;
    }

    private void ShowDrumMenu(Control anchor, int pad)
    {
        var menu = new MenuFlyout();
        for (int i = 0; i < _drumNames.Count; i++)
        {
            int note = i + 35;
            var item = new MenuItem { Header = _drumNames[i] };
            item.Click += (_, _) =>
            {
                _settings.Drums[pad] = note;
                UpdatePadTooltip(pad);
            };
            menu.Items.Add(item);
        }
        menu.ShowAt(anchor);
    }

    private void UpdatePadTooltip(int pad)
    {
        int idx = _settings.Drums[pad] - 35;
        string name = idx >= 0 && idx < _drumNames.Count ? _drumNames[idx] : $"Note {_settings.Drums[pad]}";
        ToolTip.SetTip(_drumPads[pad], name);
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
        var top = TopLevel.GetTopLevel(this);
        if (top is null) return;

        var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("SoundFont") { Patterns = new[] { "*.sf2", "*.sf3" } },
            },
        });

        var file = files.FirstOrDefault();
        if (file is null) return;

        string path = file.TryGetLocalPath() ?? file.Path.AbsolutePath;
        if (TryUseSoundFont(path)) SelectInstrument(_instrSelect);
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
            if (_noteLabels[i] is { } label) label.Text = NoteNameFor(i);
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
        foreach (Key k in _held.ToArray())
        {
            _held.Remove(k);
            int p = Array.IndexOf(PianoKeys, k);
            if (p >= 0) SetKey(p, false);
            int d = Array.IndexOf(DrumKeys, k);
            if (d >= 0) SetDrum(d, false);
        }
        for (int b = 0; b < 4; b++) _bankButtons[b].Background = BtnNormal;
    }

    // ---------- keyboard ----------

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        Key k = e.Key;
        if (k == Key.None || _held.Contains(k)) return;   // ignore auto-repeat

        bool used = false;
        if (k == Key.PageUp || k == Key.PageDown)
        {
            int dir = k == Key.PageDown ? 1 : -1;
            int next = (_instrSelect + dir + _instruments.Count) % _instruments.Count;
            SelectInstrument(next);
            used = true;
        }
        else if (k >= Key.F1 && k <= Key.F4)
        {
            int bank = k - Key.F1;
            SelectInstrument(_settings.Banks[bank]);
            _bankButtons[bank].Background = Active;
            used = true;
        }
        else if (k >= Key.F5 && k <= Key.F9)
        {
            BeatControl(k - Key.F5);
            used = true;
        }

        int p = Array.IndexOf(PianoKeys, k);
        if (p >= 0) { SetKey(p, true); used = true; }
        int d = Array.IndexOf(DrumKeys, k);
        if (d >= 0) { SetDrum(d, true); used = true; }

        if (used) { _held.Add(k); e.Handled = true; }
    }

    private void OnKeyUp(object? sender, KeyEventArgs e)
    {
        Key k = e.Key;
        if (k == Key.None) return;
        _held.Remove(k);

        if (k >= Key.F1 && k <= Key.F4)
            _bankButtons[k - Key.F1].Background = BtnNormal;

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
        var box = new TextBlock
        {
            Text = message,
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(0xE6, 0xE6, 0xEA)),
            Margin = new Thickness(20),
            MaxWidth = 460,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var ok = new Button
        {
            Content = "OK",
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(20),
            Padding = new Thickness(16, 8),
            CornerRadius = new CornerRadius(6),
            Background = BtnNormal,
            BorderBrush = BtnBorder,
            BorderThickness = new Thickness(1),
            Foreground = BtnFg,
        };

        var dialog = new Window
        {
            Title = "My MIDI",
            Width = 480,
            Height = 200,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1E)),
            Content = new Grid
            {
                RowDefinitions = { new RowDefinition(GridLength.Star), new RowDefinition(GridLength.Auto) },
                Children = { box, ok },
            },
        };
        Grid.SetRow(ok, 1);
        ok.Click += (_, _) => dialog.Close();

        await dialog.ShowDialog(this);
    }
}