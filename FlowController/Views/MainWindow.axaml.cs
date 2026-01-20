using Avalonia.Controls;
using Avalonia.Threading;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Metadata;
using Avalonia.Rendering;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Reflection.Emit;

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

        private readonly Dictionary<string, (CancellationTokenSource cts, System.Threading.Tasks.Task task)> _scan
            = new(StringComparer.OrdinalIgnoreCase);

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

        private CancellationTokenSource? _continueCts;
        private Task? _continueTask;

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

        public MainWindow()
        {
            InitializeComponent();
            StartWatcher();
            ScanCOMs();
            ContiueScanning();
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

            lock (_devicesLock)
            {
                snapshot = _devices.ToList();

                var coms = SerialPort.GetPortNames().ToList();

                snapshot = snapshot
                    .OrderBy(d =>
                    {
                        var com = Regex.Match(d, @"COM\d+", RegexOptions.IgnoreCase).Value;
                        var idx = coms.IndexOf(com);
                        return idx == -1 ? int.MaxValue : idx;
                    })
                    .ThenBy(d =>
                    {
                        var m = Regex.Match(d, @"ID\s*([A-Z])", RegexOptions.IgnoreCase);
                        return m.Success ? m.Groups[1].Value.ToUpperInvariant() : "Z";
                    })
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
                Devices.SelectedItem = _selected;
            }
            else if (snapshot.Count > 0)
            {
                Devices.SelectedItem = snapshot[0];
            }
            else
            {
                Devices.SelectedItem = null;
            }
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