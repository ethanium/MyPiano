using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
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

    private readonly MidiOut _midi = new();
    private readonly Settings _settings = Settings.Load();
    private readonly DispatcherTimer _beatTimer = new();

    private readonly List<string> _instruments = new();
    private readonly List<string> _drumNames = new();   // index 0 == MIDI note 35
    private IReadOnlyList<string> _devices = Array.Empty<string>();

    private readonly bool[] _held = new bool[256];
    private readonly Border[] _keyShapes = new Border[PianoKeys.Length];
    private readonly Button[] _drumPads = new Button[12];
    private readonly int[] _whitesBefore = new int[PianoKeys.Length];
    private int _whiteCount;
    private readonly Button[] _bankButtons;

    private int _instrSelect;
    private int _baseKey;
    private int _volume = 127;

    private int _beatSelect;
    private int _beatPtr;
    private int _beatMode = -1;   // -1 stopped, else section index
    private int _beatNext = -2;   // -2 none queued

    private static SolidColorBrush Brush(byte r, byte g, byte b) => new(Color.FromArgb(255, r, g, b));
    private static readonly SolidColorBrush Active = new(Colors.Gold);
    private static readonly SolidColorBrush Idle = Brush(0x80, 0x80, 0x80);

    public MainWindow()
    {
        InitializeComponent();
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1000, 560));

        _bankButtons = new[] { BtnBank0, BtnBank1, BtnBank2, BtnBank3 };

        LoadInstrumentNames();
        BuildPiano();
        BuildDrumPads();
        WireControls();

        // Devices
        _devices = MidiOut.GetDeviceNames();
        if (_settings.DeviceId >= _devices.Count) _settings.DeviceId = 0;
        OpenDevice(_settings.DeviceId);

        SelectBeat(0, start: false);
        SelectInstrument(_settings.Banks[0]);

        Root.AddHandler(UIElement.KeyDownEvent, new KeyEventHandler(OnKeyDown), true);
        Root.AddHandler(UIElement.KeyUpEvent, new KeyEventHandler(OnKeyUp), true);
        Root.Loaded += (_, _) => Root.Focus(FocusState.Programmatic);

        Activated += (_, e) =>
        {
            if (e.WindowActivationState == WindowActivationState.Deactivated) ReleaseAll();
        };
        Closed += (_, _) =>
        {
            _beatTimer.Stop();
            _midi.Close();
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
            var shape = new Border
            {
                Background = black ? new SolidColorBrush(Colors.Black) : new SolidColorBrush(Colors.White),
                BorderBrush = new SolidColorBrush(Colors.Black),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(0, 0, 4, 4),
                Child = new TextBlock
                {
                    Text = KeyLabel(PianoKeys[i]),
                    FontSize = 14,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(black ? Colors.White : Colors.Black),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Margin = new Thickness(0, 0, 0, 8),
                    IsHitTestVisible = false,
                },
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

    /// <summary>Scales the keyboard to fill the available space.</summary>
    private void LayoutPiano(double width, double height)
    {
        if (_whiteCount == 0 || width <= 0 || height <= 0) return;

        double whiteW = width / _whiteCount;
        double whiteH = Math.Max(100, height);
        double blackW = whiteW * 0.62;
        double blackH = whiteH * 0.6;
        double font = Math.Clamp(whiteW * 0.32, 9, 18);

        for (int i = 0; i < _keyShapes.Length; i++)
        {
            bool black = BlackSemitones.Contains(i % 12);
            var shape = _keyShapes[i];
            shape.Width = black ? blackW : whiteW;
            shape.Height = black ? blackH : whiteH;
            Canvas.SetLeft(shape, black ? _whitesBefore[i] * whiteW - blackW / 2 : _whitesBefore[i] * whiteW);
            if (shape.Child is TextBlock label) label.FontSize = font;
        }
        PianoCanvas.Width = width;
        PianoCanvas.Height = whiteH;
    }

    private void BuildDrumPads()
    {
        for (int i = 0; i < 12; i++)
        {
            int index = i;
            var pad = new Button
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Padding = new Thickness(0),
                Content = KeyLabel(DrumKeys[i]),
            };
            Grid.SetRow(pad, i / 3);
            Grid.SetColumn(pad, i % 3);
            pad.Click += (_, _) => ShowDrumMenu(pad, index);
            _drumPads[i] = pad;
            DrumGrid.Children.Add(pad);
            UpdatePadTooltip(i);
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
            ToolTipService.SetToolTip(btn, $"F{b + 1}");
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

        SldTempo.ValueChanged += (_, _) => UpdateBeatInterval();
        SldVolume.ValueChanged += (_, e) => _volume = (int)e.NewValue;

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
                    OpenDevice(_settings.DeviceId);   // original reopened the device to silence hanging notes
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
            flyout.ShowAt(BtnDevice);
        };

        _beatTimer.Tick += (_, _) => BeatTick();
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
            _settings.DeviceId = id;
            _midi.Open(id);
            BtnDevice.Content = id < _devices.Count ? _devices[id] : $"Device {id}";
        }
        catch (Exception ex)
        {
            BtnDevice.Content = "(none)";
            ShowError(ex.Message);
        }
    }

    private void SelectInstrument(int index)
    {
        _instrSelect = index;
        _midi.ChangeInstrument(0, index);
        BtnInstrument.Content = _instruments[index];
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
            shape.Background = Brush(black ? (byte)100 : (byte)200, black ? (byte)100 : (byte)200, black ? (byte)100 : (byte)200);
            _midi.NoteOn(0, i + _baseKey + 48, _volume);
        }
        else
        {
            shape.Background = new SolidColorBrush(black ? Colors.Black : Colors.White);
            _midi.NoteOff(0, i + _baseKey + 48);
        }
    }

    private void SetDrum(int i, bool down)
    {
        if (down)
        {
            _drumPads[i].Background = Idle;
            _midi.NoteOn(9, _settings.Drums[i], 127);
        }
        else
        {
            _drumPads[i].ClearValue(Control.BackgroundProperty);
            _midi.NoteOff(9, _settings.Drums[i]);
        }
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
            _bankButtons[k - (int)VirtualKey.F1].ClearValue(Control.BackgroundProperty);

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
            Title = "MyPiano",
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = Content.XamlRoot,
        }.ShowAsync();
    }
}
