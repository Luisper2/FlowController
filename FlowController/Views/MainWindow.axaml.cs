using Avalonia;
using System.IO;
using System;
using System.Text;
using Avalonia.Input;
using System.IO.Ports;
using Avalonia.Metadata;
using System.Threading;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.Rendering;
using System.Diagnostics;
using Avalonia.VisualTree;
using System.Globalization;
using System.Reflection.Emit;
using System.Threading.Tasks;
using Avalonia.Interactivity;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Net.Mime.MediaTypeNames;

#if WINDOWS
using System.Management;
#endif

// Disable windows warning for Windows Watcher
#pragma warning disable CA1416

internal sealed class PortSession : IDisposable
{
    public string Com { get; }
    public SerialPort Port { get; }

    private readonly SemaphoreSlim _gate = new(1, 1);
    private bool _disposed;

    public PortSession(string com, int baud = 38400, int timeout = 400)
    {
        Com = com;

        Port = new SerialPort(com)
        {
            BaudRate = baud,
            DataBits = 8,
            Parity = Parity.None,
            StopBits = StopBits.One,
            Encoding = Encoding.ASCII,
            NewLine = "\r",
            ReadTimeout = timeout,
            WriteTimeout = timeout
        };

        Port.Open();
        try { Port.DiscardInBuffer(); } catch { }
        try { Port.DiscardOutBuffer(); } catch { }
    }

    public async Task<T> WithPort<T>(Func<SerialPort, Task<T>> action, CancellationToken ct)
    {
        ThrowIfDisposed();

        await _gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();

            if (!Port.IsOpen)
                Port.Open();

            return await action(Port).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    public Task<(bool ok, string response)> SendCommand(string command, int timeout, CancellationToken ct)
    {
        return WithPort(sp =>
        {
            try
            {
                sp.WriteTimeout = Math.Min(100, timeout);
                sp.ReadTimeout = Math.Min(300, timeout);

                try { sp.DiscardInBuffer(); } catch { }
                try { sp.DiscardOutBuffer(); } catch { }

                sp.Write(command);

                var sw = Stopwatch.StartNew();
                var sb = new StringBuilder();

                while (sw.ElapsedMilliseconds < timeout)
                {
                    ct.ThrowIfCancellationRequested();

                    try
                    {
                        int ch = sp.ReadChar();
                        if (ch < 0) continue;

                        char c = (char)ch;

                        if (c == '\r' || c == '\n')
                        {
                            if (sb.Length == 0) continue;
                            break;
                        }

                        sb.Append(c);
                    }
                    catch (TimeoutException) {}
                    catch
                    {
                        break;
                    }
                }

                var resp = sb.ToString().Trim();
                return Task.FromResult(string.IsNullOrWhiteSpace(resp)
                    ? (false, string.Empty)
                    : (true, resp));
            }
            catch
            {
                return Task.FromResult((false, string.Empty));
            }
        }, ct);
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(PortSession));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try { Port.Close(); } catch { }
        try { Port.Dispose(); } catch { }
        try { _gate.Dispose(); } catch { }
    }
}

namespace FlowController.Views
{
    public partial class MainWindow : Window
    {
        private readonly HashSet<string> _coms = new(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, (CancellationTokenSource cts, System.Threading.Tasks.Task task)> _scan = new(StringComparer.OrdinalIgnoreCase);

        private ManagementEventWatcher? _watcher;
        private FileSystemWatcher? _devWatcher;

        private DateTime _lastScanReq = DateTime.MinValue;
        private readonly TimeSpan _debounce = TimeSpan.FromMilliseconds(250);

        private readonly Dictionary<string, PortSession> _ports = new(StringComparer.OrdinalIgnoreCase);
        private readonly object _portsLock = new();

        private List<string> _devices = new();
        private readonly object _devicesLock = new();

        private readonly Dictionary<string, DateTime> _lastSeen = new(StringComparer.OrdinalIgnoreCase);
        private readonly object _seenLock = new();

        private string _selected = string.Empty;
        private CancellationTokenSource? _selectedCts;
        private Task? _selectedTask;
        private float _setpoint = float.NaN;

        private CancellationTokenSource? _spRepeatCts;
        private Key? _spRepeatKey;
        private readonly Stopwatch _spHoldSw = new();
        private bool _setpointDirty = false;
        private CancellationTokenSource? _sendSpDebounceCts;

        private const float SpBaseStepSccm = 1f;

        private CancellationTokenSource? _continueCts;
        private Task? _continueTask;

        private CancellationTokenSource? _idRepeatCts;
        private Key? _idRepeatKey;
        private readonly Stopwatch _idHoldSw = new();

        private int _scanQueued = 0;
        
        private static bool IsSerialDevice(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;

            return name.StartsWith("ttyUSB", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ttyACM", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("rfcomm", StringComparison.OrdinalIgnoreCase);
        }
        
        private static bool IsValidSerial(string serial)
        {
            if (string.IsNullOrWhiteSpace(serial))
                return false;

            var regex = new Regex(@"^[A-Z]\s[A-Z]\d{6}$");

            return regex.IsMatch(serial.Trim());
        }

        static void Change(TextBlock tx, string val, float eps = 0.001f, string unit = "")
        {
            float.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var value);

            var valueText = (tx.Text ?? "-");
            float.TryParse(valueText, NumberStyles.Float, CultureInfo.InvariantCulture, out float valueUI);

            if (string.IsNullOrWhiteSpace(valueText) || (Math.Abs(value - valueUI) > eps) || value == 0) tx.Text = value.ToString("0.###", CultureInfo.InvariantCulture) + unit;
        }

        static double toSCCM(double value, string unit)
        {
            var conversions = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
            {
                { "SCCM", 1.0 },
                { "SmL/s", 60.0 },
                { "SmL/m", 1.0 },
                { "SLPM", 1000.0 },
                { "SL/h", 16.6667 },
                { "SCCS", 1.0 / 60.0 },
                { "Sm3/h", 1000.0 / 60.0 },
                { "SCIM", 0.0610237 },
                { "SCFM", 35.3147 * 1000.0 },
                { "SCFH", 35.3147 * 1000.0 / 60.0 },
                { "SCFD", 35.3147 * 1000.0 / 60.0 / 24.0 }
            };

            if (!conversions.TryGetValue(unit, out var factor))
                return 0.0;

            return value * factor;
        }

        static double fromSCCM(double value, string unit)
        {
            var conversions = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
            {
                { "SCCM", 1.0 },
                { "SmL/s", 1.0 / 60.0 },
                { "SmL/m", 1.0 },
                { "SLPM", 1.0 / 1000.0 },
                { "SL/h", 1.0 / 16.6667 },
                { "SCCS", 60.0 },
                { "Sm3/h", 60.0 / 1000.0 },
                { "SCIM", 1.0 / 0.0610237 },
                { "SCFM", 1.0 / (35.3147 * 1000.0) },
                { "SCFH", 1.0 / (35.3147 * 1000.0 / 60.0) },
                { "SCFD", 1.0 / (35.3147 * 1000.0 / 60.0 / 24.0) }
            };

            if (!conversions.TryGetValue(unit, out var factor))
                return 0.0;

            var converted = value * factor;

            return Math.Abs(converted % 1) < 1e-9
                ? Math.Truncate(converted)
                : Math.Round(converted, 3, MidpointRounding.AwayFromZero);
        }

        public MainWindow()
        {
            InitializeComponent();

            Dispatcher.UIThread.Post(() =>
            {
                Devices.Focus();
                SelectedDeviceItem();
            });

            StartWatcher();
            ScanCOMs();
            ContiueScanning();

            Devices.SelectionChanged += Devices_SelectionChanged;

            Dispatcher.UIThread.Post(() =>
            {
                _selected = Devices.SelectedItem as string ?? string.Empty;
                PollSelected(_selected);
            });

            Devices.SelectionChanged += (_, __) => {
                _setpoint = float.NaN;
            };

            Devices.GotFocus += (_, __) =>
            {
                DevicesLabel.Foreground = Avalonia.Media.Brushes.Red;
            };

            Devices.LostFocus += (_, __) =>
            {
                DevicesLabel.Foreground = Avalonia.Media.Brushes.Black;
            };

            Setpoint.AddHandler(InputElement.KeyDownEvent, (_, e) =>
            {
                if (!Setpoint.IsFocused || (e.Key != Key.Up && e.Key != Key.Down)) return;

                if (float.IsNaN(_setpoint)) return;

                var unit = (Units.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(unit) || unit == "-") return;

                e.Handled = true;

                if (_spRepeatCts != null && _spRepeatKey == e.Key) return;

                StopSetpointRepeat();

                _spRepeatKey = e.Key;
                _spRepeatCts = new CancellationTokenSource();
                var token = _spRepeatCts.Token;

                _spHoldSw.Restart();

                _ = Task.Run(async () =>
                {
                    try
                    {
                        ApplySetpointStep(e.Key, unit, _spHoldSw.ElapsedMilliseconds);

                        const int holdThresholdMs = 250;
                        await Task.Delay(holdThresholdMs, token);

                        while (!token.IsCancellationRequested)
                        {
                            var ms = _spHoldSw.ElapsedMilliseconds;

                            int delay = ms < 800 ? 90
                                : ms < 1500 ? 75
                                : 60;

                            await Task.Delay(delay, token);

                            ApplySetpointStep(e.Key, unit, ms);
                        }
                    }
                    catch (OperationCanceledException) { }
                }, token);
            }, RoutingStrategies.Tunnel, handledEventsToo: true);

            Setpoint.AddHandler(InputElement.KeyUpEvent, (_, e) =>
            {
                if (e.Key == Key.Up || e.Key == Key.Down)
                {
                    StopSetpointRepeat();
                    e.Handled = true;

                    SetSetpoint();
                }
            }, RoutingStrategies.Tunnel, handledEventsToo: true);

            Setpoint.GotFocus += (_, __) =>
            {
                SetpointLabel.Foreground = Avalonia.Media.Brushes.Red;
            };

            Setpoint.LostFocus += (_, __) =>
            {
                SetpointLabel.Foreground = Avalonia.Media.Brushes.Black;
            };

            Setup.AddHandler(InputElement.KeyDownEvent, (_, e) =>
            {
                if (!Setup.IsFocused || e.Key != Key.Down) return;

                e.Handled = true;

                if (Setup.Content is string text && text == "Setup")
                {
                    // Change Visibility
                    Main.IsVisible = false;
                    SetupMenu.IsVisible = true;

                    // Modify TabIndex
                    Setpoint.TabIndex = -1;

                    ID.TabIndex = 1;
                    ChangeButton.TabIndex = 2;
                    Setup.TabIndex = 3;

                    Setup.Content = "Back";

                    return;
                }

                // Change Visibility
                Main.IsVisible = true;
                SetupMenu.IsVisible = false;

                // Modify TabIndex
                ID.TabIndex = -1;
                ChangeButton.TabIndex = -1;

                Setpoint.TabIndex = 1;
                Setup.TabIndex = 2;

                Setup.Content = "Setup";
            }, RoutingStrategies.Tunnel);

            Setup.GotFocus += (_, __) =>
            {
                Setup.Foreground = Avalonia.Media.Brushes.Red;
                Setup.Background = Avalonia.Media.Brushes.White;
            };

            Setup.LostFocus += (_, __) =>
            {
                Setup.Foreground = Avalonia.Media.Brushes.Black;
                Setup.Background = Avalonia.Media.Brushes.LightGray;
            };

            ID.AddHandler(InputElement.KeyDownEvent, (_, e) =>
            {
                if (!ID.IsFocused) return;
                if (e.Key != Key.Up && e.Key != Key.Down) return;

                e.Handled = true;

                if (_idRepeatCts != null && _idRepeatKey == e.Key) return;

                StopIdRepeat();

                _idRepeatKey = e.Key;
                _idRepeatCts = new CancellationTokenSource();
                var token = _idRepeatCts.Token;

                _idHoldSw.Restart();

                _ = Task.Run(async () =>
                {
                    try
                    {
                        Dispatcher.UIThread.Post(() => IDStep(e.Key));

                        const int holdThresholdMs = 250;
                        await Task.Delay(holdThresholdMs, token);

                        while (!token.IsCancellationRequested)
                        {
                            var ms = _idHoldSw.ElapsedMilliseconds;

                            int delay = ms < 800 ? 90
                                : ms < 1500 ? 75
                                : 60;

                            await Task.Delay(delay, token);

                            Dispatcher.UIThread.Post(() => IDStep(e.Key));
                        }
                    }
                    catch (OperationCanceledException) { }
                }, token);
            }, RoutingStrategies.Tunnel, handledEventsToo: true);

            ID.AddHandler(InputElement.KeyUpEvent, (_, e) =>
            {
                if (e.Key == Key.Up || e.Key == Key.Down)
                {
                    StopIdRepeat();
                    e.Handled = true;
                }
            }, RoutingStrategies.Tunnel, handledEventsToo: true);

            ID.GotFocus += (_, __) =>
            {
                ID.Foreground = Avalonia.Media.Brushes.Red;
            };

            ID.LostFocus += (_, __) =>
            {
                ID.Foreground = Avalonia.Media.Brushes.Black;
            };

            ChangeButton.AddHandler(InputElement.KeyDownEvent, async (_, e) =>
            {
                if (!ChangeButton.IsFocused || e.Key != Key.Down) return;

                e.Handled = true;

                var selectedLabel = _selected;
                if (string.IsNullOrWhiteSpace(selectedLabel)) return;

                var baseId = selectedLabel.Split(" - ")[0].Trim();
                var parts = baseId.Split('/');
                if (parts.Length != 2) return;

                var com = parts[0].Trim();
                var oldId = parts[1].Trim();

                var newId = (ID.Text ?? "").Trim().ToUpperInvariant();
                if (newId.Length != 1 || newId[0] < 'A' || newId[0] > 'Z') return;
                if (string.Equals(oldId, newId, StringComparison.OrdinalIgnoreCase)) return;

                try { _selectedCts?.Cancel(); } catch { }
                _selectedCts?.Dispose();
                _selectedCts = null;
                _selectedTask = null;

                try
                {
                    using var cmdCts = new CancellationTokenSource(800);
                    var session = GetPort(com);

                    var cmd = $"{oldId}@={newId}\r";
                    var (ok, resp) = await session.SendCommand(cmd, timeout: 600, ct: cmdCts.Token);

                    if (!ok)
                    {
                        System.Diagnostics.Debug.WriteLine($"Rename failed: {com}/{oldId} -> {newId}. resp='{resp}'");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Rename exception: {ex}");
                    return;
                }

                var oldKey = $"{com}/{oldId}";
                var newKey = $"{com}/{newId}";

                lock (_devicesLock)
                {
                    int idx = _devices.FindIndex(d =>
                        d.Equals(oldKey, StringComparison.OrdinalIgnoreCase) ||
                        d.StartsWith(oldKey + " -", StringComparison.OrdinalIgnoreCase));

                    if (idx >= 0)
                    {
                        bool hadError = _devices[idx].Contains(" - Error", StringComparison.OrdinalIgnoreCase);
                        _devices[idx] = hadError ? $"{newKey} - Error" : newKey;
                    }
                    else
                    {
                        _devices.Add(newKey);
                    }
                }

                lock (_seenLock)
                {
                    _lastSeen.Remove(oldKey);
                    _lastSeen[newKey] = DateTime.UtcNow;
                }

                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    DisplayDevices();

                    Devices.SelectedItem = newKey;
                    _selected = newKey;

                    ID.Text = newId;

                    PollSelected(_selected);

                    _setpoint = float.NaN;
                });
            }, RoutingStrategies.Tunnel, handledEventsToo: true);

            ChangeButton.GotFocus += (_, __) =>
            {
                ChangeButton.Foreground = Avalonia.Media.Brushes.Red;
                ChangeButton.Background = Avalonia.Media.Brushes.White;
            };

            ChangeButton.LostFocus += (_, __) =>
            {
                ChangeButton.Foreground = Avalonia.Media.Brushes.Black;
                ChangeButton.Background = Avalonia.Media.Brushes.LightGray;
            };
        }

        private void StartWatcher()
        {
            if (!OperatingSystem.IsWindows() && !OperatingSystem.IsLinux()) return;

            if (OperatingSystem.IsWindows())
            {
                try
                {
                    _watcher = new ManagementEventWatcher(
                        new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 OR EventType = 3")
                    );

                    _watcher.EventArrived += (_, __) =>
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            var now = DateTime.UtcNow;
                            if (now - _lastScanReq < _debounce) return;
                            _lastScanReq = now;
                            
                            ScanCOMs();
                        });
                    };

                    _watcher.Start();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Device had not been detected. Error: {ex}");
                }

                return;
            }

            if (OperatingSystem.IsLinux())
            {
                try
                {
                    _devWatcher = new FileSystemWatcher("/dev")
                    {
                        IncludeSubdirectories = false,
                        NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
                    };

                    _devWatcher.Created += (_, e) =>
                    {
                        if (IsSerialDevice(e.Name))
                            QueueScanOnUI();
                    };

                    _devWatcher.Deleted += (_, e) =>
                    {
                        if (IsSerialDevice(e.Name))
                            QueueScanOnUI();
                    };

                    _devWatcher.Renamed += (_, e) =>
                    {
                        if (IsSerialDevice(e.OldName) || IsSerialDevice(e.Name))
                            QueueScanOnUI();
                    };

                    _devWatcher.EnableRaisingEvents = true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Linux /dev watcher error: {ex}");
                }

                return;
            }
        }

        private void QueueScanOnUI()
        {
            if (Interlocked.Exchange(ref _scanQueued, 1) == 1) return;

            Dispatcher.UIThread.Post(() =>
            {
                try
                {
                    var now = DateTime.UtcNow;
                    if (now - _lastScanReq < _debounce) return;
                    _lastScanReq = now;

                    ScanCOMs();
                }
                finally {
                    Interlocked.Exchange(ref _scanQueued, 0);
                }
            });
        }

        private void ScanCOMs()
        {
            try
            {
                var current = SerialPort.GetPortNames().ToHashSet(StringComparer.OrdinalIgnoreCase);

                foreach (var com in current)
                {
                    if (_coms.Add(com))
                    {
                        _ = GetPort(com);
                        ScanDevices(com);
                    }
                }

                foreach (var com in _coms.Where(c => !current.Contains(c)).ToList())
                {
                    _coms.Remove(com);
                    StopScan(com);
                    DisposePort(com);

                    RemoveDevicesFromCOM(com);
                }

                DisplayDevices();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ScanCOMs error: {ex}");
            }
        }

        private void ScanDevices(string com, int timeout = 100)
        {
            async Task ScanIDs(PortSession session, string com, CancellationTokenSource cts)
            {
                for (char id = 'A'; id <= 'Z'; id++)
                {
                    cts.Token.ThrowIfCancellationRequested();

                    string deviceId = $"{com}/{id}";

                    int idx;

                    lock (_devicesLock)
                    {
                        idx = _devices.FindIndex(d =>
                            d.Equals(deviceId, StringComparison.OrdinalIgnoreCase) ||
                            d.StartsWith(deviceId + " -", StringComparison.OrdinalIgnoreCase));
                    }

                    if (idx >= 0) continue;

                    var (ok, response) = await session.SendCommand($"{id}SN\r", timeout: timeout, ct: cts.Token);

                    if (!ok || string.IsNullOrWhiteSpace(response))
                        continue;

                    lock (_seenLock) _lastSeen[deviceId] = DateTime.UtcNow;

                    lock (_devicesLock)
                    {
                        idx = _devices.FindIndex(d =>
                            d.Equals(deviceId, StringComparison.OrdinalIgnoreCase) ||
                            d.StartsWith(deviceId + " -", StringComparison.OrdinalIgnoreCase));

                        if (idx >= 0) continue;

                        bool isValid = IsValidSerial(response);

                        string label = isValid ? deviceId : $"{deviceId} - Error";

                        _devices.Add(label);
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    var now = DateTime.UtcNow;
                    
                    if (now - _lastScanReq < _debounce) return;
                    _lastScanReq = now;

                    DisplayDevices();
                });
            }

            if (_scan.ContainsKey(com)) return;

            var cts = new CancellationTokenSource();
            _scan[com] = (cts, Task.CompletedTask);

            var task = Task.Run(async () =>
            {
                try
                {
                    var session = GetPort(com);

                    while (!cts.Token.IsCancellationRequested)
                    {
                        var (ok, response) = await session.SendCommand($"*LSS S\r", timeout: timeout, ct: cts.Token);

                        await ScanIDs(session, com, cts);
                        await Task.Delay(TimeSpan.FromSeconds(15), cts.Token);
                    }
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Scan failed ({com}): {ex.Message}");
                }
            }, cts.Token);

            _scan[com] = (cts, task);
        }

        private void ContiueScanning()
        {
            if (_continueCts != null) return;

            _continueCts = new CancellationTokenSource();
            var token = _continueCts.Token;

            TimeSpan _alive = TimeSpan.FromSeconds(3);
            TimeSpan _tick = TimeSpan.FromMilliseconds(750);

            _continueTask = Task.Run(async () => {
                while (!token.IsCancellationRequested) {
                    List<string> snapshot, comList;

                    lock (_devicesLock) snapshot = _devices.ToList();

                    if (snapshot.Count > 0)
                    {
                        comList = snapshot
                            .Select(d => d.Split("/")[0].Trim())
                            .Where(d => !string.IsNullOrWhiteSpace(d))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        bool _refresh = false;

                        var tasks = comList.Select(com =>
                            Task.Run(async () => {
                                List<string> devicesCom = snapshot
                                    .Where(d => d.StartsWith(com + "/", StringComparison.OrdinalIgnoreCase))
                                    .ToList();

                                if (devicesCom.Count == 0) return;

                                var session = GetPort(com);

                                foreach (var d in devicesCom)
                                {
                                    var id = d.Split(" - ")[0].Split("/")[1].Trim();
                                    string deviceId = $"{com}/{id}";

                                    bool needScan;
                                    var now = DateTime.UtcNow;

                                    lock (_seenLock)
                                    {
                                        needScan = !_lastSeen.TryGetValue(deviceId, out var last) || (now - last) > _alive;
                                    }

                                    if (!needScan) continue;

                                    var (ok, response) = await session.SendCommand($"{id}SN\r", 250, token);

                                    if (!ok) continue;
                                    
                                    bool _valid = IsValidSerial(response);

                                    lock (_devicesLock)
                                    {
                                        if (response == "") {
                                            _refresh = true;

                                            lock (_seenLock) _lastSeen.Remove(deviceId);
                                            _devices.Remove(d);

                                            continue;
                                        }

                                        string label = _valid ? deviceId : $"{deviceId} - Error";

                                        _refresh = true;

                                        int idx = _devices.FindIndex(dd =>
                                            dd.Equals(deviceId, StringComparison.OrdinalIgnoreCase) ||
                                            dd.StartsWith(deviceId + " -", StringComparison.OrdinalIgnoreCase));

                                        lock (_seenLock) _lastSeen[deviceId] = now;
                                        _devices[idx] = label;
                                    }
                                }
                            }, token)
                        );

                        await Task.WhenAll(tasks);

                        if (_refresh) Dispatcher.UIThread.Post(DisplayDevices);
                    }

                    await Task.Delay(_tick, token);
                }
            }, token);
        }

        private void Devices_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            _selected = Devices.SelectedItem as string ?? string.Empty;

            var label = _selected?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(label))
            {
                ID.Text = "-";
                return;
            }

            var baseId = label.Split(" - ", StringSplitOptions.RemoveEmptyEntries)[0].Trim();

            var parts = baseId.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2)
            {
                ID.Text = "-";
                return;
            }

            var id = parts[1].Trim();
            if (id.Length != 1 || id[0] < 'A' || id[0] > 'Z')
            {
                ID.Text = "-";
                return;
            }

            ID.Text = id;

            PollSelected(label);
        }

        private void PollSelected(string? selectedLabel)
        {
            try { _selectedCts?.Cancel(); } catch {}
            
            _selectedCts?.Dispose();
            _selectedCts = new CancellationTokenSource();

            var token = _selectedCts.Token;

            if (string.IsNullOrWhiteSpace(selectedLabel)) return;

            var baseId = selectedLabel.Split(" - ")[0].Trim();

            var parts = baseId.Split('/');
            if (parts.Length != 2) return;

            var com = parts[0].Trim();
            var id = parts[1].Trim();

            _selectedTask = Task.Run(async () =>
            {
                var tick = TimeSpan.FromMilliseconds(250);
                var session = GetPort(com);

                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        var (ok1, resp1) = await session.SendCommand($"{id}SN\r", 100, token);
                        var (ok2, resp2) = await session.SendCommand($"{id}\r", 100, token);
                        var (ok3, resp3) = await session.SendCommand($"{id}FPF 0\r", 100, token);
                        var (ok4, resp4) = await session.SendCommand($"{id}FPF 1\r", 100, token);
                        var (ok5, resp5) = await session.SendCommand($"{id}DV 2\r", 100, token);

                        bool _valid = IsValidSerial(resp1);

                        if (!ok1 || !_valid)
                        {
                            await Dispatcher.UIThread.InvokeAsync(() =>
                            {
                                Flow.Text = "-";
                                Units.Text = "-";
                                Range.Text = "-";
                                Setpoint.Text = "-";
                                Temperature.Text = "-";
                                Drive.Text = "-";
                                Volume.Text = "-";
                                Multiple.Opacity = 100;
                                TOV.Opacity = 0;
                                MOV.Opacity = 0;
                                OVR.Opacity = 0;
                                HLD.Opacity = 0;
                                VTM.Opacity = 0;
                            });
                        }
                        
                        if (ok1 || (ok2 || ok3 || ok4 || ok5))
                        {
                            await Dispatcher.UIThread.InvokeAsync(() =>
                            {
                                float eps = 0.001f;

                                Multiple.Opacity = 0;

                                // Flow, Setpoint, Temperature, Volume and Drive
                                var parts = resp2.Split(" ", StringSplitOptions.RemoveEmptyEntries);

                                float.TryParse(resp5.Split(" ")[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var setpointValue);

                                Change(Temperature, parts[1], eps);
                                Change(Flow, parts[2], eps);
                                Change(Volume, parts[3], eps, " " + resp4.Split(" ", StringSplitOptions.RemoveEmptyEntries)[2].Substring(1).Trim() ?? "");
                                Change(Drive, parts[5], eps);

                                if (_setpoint is float.NaN)
                                {
                                    _setpoint = setpointValue;
                                    Change(Setpoint, resp5.Split(" ")[1].Trim(), eps);
                                }

                                // Errors
                                TOV.Opacity = parts.Contains("TOV") && TOV.Opacity == 0 ? 100 : 0;
                                MOV.Opacity = parts.Contains("MOV") && TOV.Opacity == 0 ? 100 : 0;
                                OVR.Opacity = parts.Contains("OVR") && TOV.Opacity == 0 ? 100 : 0;
                                HLD.Opacity = parts.Contains("HLD") && TOV.Opacity == 0 ? 100 : 0;
                                VTM.Opacity = parts.Contains("VTM") && TOV.Opacity == 0 ? 100 : 0;

                                // Units and Range
                                parts = resp3.Split(" ", StringSplitOptions.RemoveEmptyEntries);

                                if (Units.Text != parts[2]) Units.Text = parts[2];

                                float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var value);
                                float valueUI = float.NaN;

                                var text = (Range.Text ?? "").Trim();

                                if (text != "-" && text.Contains(" - "))
                                {
                                    var partsRange = text.Split(" - ", StringSplitOptions.RemoveEmptyEntries);
                                    
                                    if (partsRange.Length == 2) float.TryParse(partsRange[1], NumberStyles.Float, CultureInfo.InvariantCulture, out valueUI);
                                }

                                if (float.IsNaN(valueUI) || Math.Abs(value - valueUI) > eps) Range.Text = $"0 - {value.ToString("0.###", CultureInfo.InvariantCulture)}";
                            });
                        }
                    }
                    catch (OperationCanceledException) { break; }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Selected poll error ({com}/{id}): {ex.Message}");
                    }

                    await Task.Delay(tick, token);
                }
            }, token);
        }

        private void StopSetpointRepeat()
        {
            try { _spRepeatCts?.Cancel(); } catch { }
            _spRepeatCts?.Dispose();
            _spRepeatCts = null;
            _spRepeatKey = null;
        }

        private void ApplySetpointStep(Key key, string unit, long heldMs)
        {
            var rangeText = Dispatcher.UIThread.InvokeAsync(() => Range.Text ?? "").GetAwaiter().GetResult();
            float.TryParse(rangeText.Split(" - ")[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var maxValue);

            if (float.IsNaN(_setpoint) || string.IsNullOrWhiteSpace(rangeText) || !rangeText.Contains(" - ")) return;
            
            _setpointDirty = true;

            double currentInUnit = _setpoint;
            double currentSccm = toSCCM(currentInUnit, unit);

            double stepSccm = heldMs < 500 ? SpBaseStepSccm
                : heldMs < 1200 ? SpBaseStepSccm * 5
                : heldMs < 2500 ? SpBaseStepSccm * 10
                : SpBaseStepSccm * 25;

            currentSccm += (key == Key.Up) ? stepSccm : -stepSccm;

            double newInUnit = fromSCCM(currentSccm, unit);
            
            if (newInUnit > maxValue)
            {
                _setpoint = 0;
            }
            else if (newInUnit < 0)
            {
                _setpoint = maxValue;
            }
            else
            {
                _setpoint = (float)newInUnit;
            }

            Dispatcher.UIThread.Post(() =>
            {
                Setpoint.Text = _setpoint.ToString("0.###", CultureInfo.InvariantCulture);
            });
        }

        private void SetSetpoint()
        {
            if (!_setpointDirty) return;
            if (float.IsNaN(_setpoint)) return;

            var label = _selected;
            if (string.IsNullOrWhiteSpace(label)) return;

            var baseId = label.Split(" - ")[0].Trim();
            var parts = baseId.Split('/');
            if (parts.Length != 2) return;

            var com = parts[0].Trim();
            var id = parts[1].Trim();

            var unit = (Units.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(unit) || unit == "-") return;

            try { _sendSpDebounceCts?.Cancel(); } catch { }
            _sendSpDebounceCts?.Dispose();
            _sendSpDebounceCts = new CancellationTokenSource();
            var token = _sendSpDebounceCts.Token;

            _ = Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(150, token);
                    if (token.IsCancellationRequested) return;

                    var session = GetPort(com);

                    var (ok, resp) = await session.SendCommand($"{id}S {_setpoint.ToString("0.###", CultureInfo.InvariantCulture)}\r", 100, token);

                    if (ok)
                    {
                        _setpointDirty = false;
                    }
                }
                catch { }
            }, token);
        }

        private void IDStep(Key key)
        {
            var text = (ID.Text ?? "A").Trim().ToUpperInvariant();
            char current = (text.Length > 0 && text[0] >= 'A' && text[0] <= 'Z') ? text[0] : 'A';

            int idx = current - 'A';
            idx += (key == Key.Up) ? -1 : +1;

            if (idx < 0) idx = 25;
            if (idx > 25) idx = 0;

            char next = (char)('A' + idx);
            ID.Text = next.ToString();
        }
        
        private void StopIdRepeat()
        {
            try { _idRepeatCts?.Cancel(); } catch { }
            _idRepeatCts?.Dispose();
            _idRepeatCts = null;
            _idRepeatKey = null;
        }

        private PortSession GetPort(string com)
        {
            lock (_portsLock)
            {
                if (_ports.TryGetValue(com, out var s))
                    return s;

                var session = new PortSession(com, baud: 38400, timeout: 400);
                _ports[com] = session;
                return session;
            }
        }

        private void DisposePort(string com)
        {
            lock (_portsLock)
            {
                if (_ports.Remove(com, out var s))
                    s.Dispose();
            }
        }

        private void StopScan(string com)
        {
            if (!_scan.TryGetValue(com, out var w)) return;

            _scan.Remove(com);

            try { w.cts.Cancel(); } catch {}

            _ = w.task.ContinueWith(_ =>
            {
                try { w.cts.Dispose(); } catch { }
            });
        }
        
        private void DisplayDevices()
        {
            List<string> snapshot;
            string newSelected;

            lock (_devicesLock)
            {
                snapshot = _devices.ToList();

                var coms = SerialPort.GetPortNames().ToList();

                snapshot = snapshot
                    .Select(d =>
                    {
                        var baseId = d.Split(" - ")[0].Trim();
                        var parts = baseId.Split('/');
                        return new
                        {
                            Original = d,
                            Com = parts.Length > 0 ? parts[0] : "",
                            Id = parts.Length > 1 ? parts[1] : "Z"
                        };
                    })
                    .OrderBy(x => 
                    {
                        lock (_seenLock)
                        {
                            var times = _lastSeen
                                .Where(kv => kv.Key.StartsWith(x.Com + "/", StringComparison.OrdinalIgnoreCase))
                                .Select(kv => kv.Value);

                            return times.Any() ? times.Min() : DateTime.MaxValue;
                        }
                    })
                    .ThenBy(x => x.Id, StringComparer.OrdinalIgnoreCase)
                    .Select(x => x.Original)
                    .ToList();
            }

            var snapshotSet = new HashSet<string>(snapshot);

            for (int i = Devices.Items.Count - 1; i >= 0; i--)
            {
                if (Devices.Items[i] is string existing && !snapshotSet.Contains(existing))
                    Devices.Items.RemoveAt(i);
            }

            var currentSet = new HashSet<string>(Devices.Items.Cast<object>()
                                                           .OfType<string>());

            foreach (var d in snapshot)
            {
                if (!currentSet.Contains(d))
                    Devices.Items.Add(d);
            }

            for (int targetIndex = 0; targetIndex < snapshot.Count; targetIndex++)
            {
                var desired = snapshot[targetIndex];

                int currentIndex = -1;
                for (int i = 0; i < Devices.Items.Count; i++)
                {
                    if (Devices.Items[i] is string s && s == desired)
                    {
                        currentIndex = i;
                        break;
                    }
                }

                if (currentIndex == -1) continue;

                if (currentIndex != targetIndex)
                {
                    Devices.Items.RemoveAt(currentIndex);
                    Devices.Items.Insert(targetIndex, desired);
                }
            }

            if (!string.IsNullOrEmpty(_selected) && snapshotSet.Contains(_selected))
            {
                newSelected = _selected;
            }
            else if (snapshot.Count > 0)
            {
                newSelected = snapshot[0];
            }
            else
            {
                newSelected = string.Empty;
            }

            var currentUiSel = Devices.SelectedItem as string ?? string.Empty;

            if (!string.Equals(currentUiSel, newSelected, StringComparison.OrdinalIgnoreCase))
            {
                Devices.SelectedItem = string.IsNullOrEmpty(newSelected) ? null : newSelected;
                _selected = newSelected;

                PollSelected(_selected);
                SelectedDeviceItem();
            }
        }

        private void SelectedDeviceItem()
        {
            if (!Devices.IsFocused) return;

            Dispatcher.UIThread.Post(() =>
            {
                if (Devices.SelectedIndex < 0 && Devices.ItemCount > 0)
                    Devices.SelectedIndex = 0;

                var idx = Devices.SelectedIndex;
                if (idx < 0) return;
            }, DispatcherPriority.Background);
        }

        private void RemoveDevicesFromCOM(string com)
        {
            lock (_devicesLock)
            {
                _devices.RemoveAll(d => d.StartsWith(com + "/", StringComparison.OrdinalIgnoreCase));
            }
        }
        
        protected override void OnClosed(EventArgs e)
        {
            #if WINDOWS
            if (_watcher is not null)
            {
                try
                {
                    _watcher.Stop();
                    _watcher.Dispose();
                }
                catch {}

                _watcher = null;
            }
            #endif

            if (_devWatcher is not null)
            {
                try
                {
                    _devWatcher.EnableRaisingEvents = false;
                    _devWatcher.Dispose();
                }
                catch {}

                _devWatcher = null;
            }

            _continueCts?.Cancel();

            try { _continueTask?.Wait(200); } catch {}

            _continueCts?.Dispose();

            foreach (var com in _scan.Keys.ToList()) StopScan(com);

            _scan.Clear();

            base.OnClosed(e);
        }
    }
}