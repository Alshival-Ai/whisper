using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Shapes;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using IOPath = System.IO.Path;

namespace Whisper
{
    public sealed partial class MainWindow : Window
    {
        private const double WideLayoutBreakpoint = 980;
        private const string DefaultWhisperModel = "openai/whisper-medium";
        private readonly Windows.UI.Color[] _teamAccentColors =
        {
            ColorHelper.FromArgb(0xFF, 0x5D, 0xA8, 0xFF),
            ColorHelper.FromArgb(0xFF, 0x7A, 0xB6, 0xF5),
            ColorHelper.FromArgb(0xFF, 0x7A, 0xCB, 0x78),
            ColorHelper.FromArgb(0xFF, 0xA2, 0x9B, 0xFE),
            ColorHelper.FromArgb(0xFF, 0xF2, 0xA6, 0x5A),
            ColorHelper.FromArgb(0xFF, 0x4D, 0xD0, 0xC8),
        };

        private int _nextTeamColorIndex = 3;
        private int _nextCreatedTeamNumber = 1;
        private bool _isInNotesMode;
        private bool _isInSettingsMode;
        private bool _isSidebarCollapsedByUser;
        private bool _isCaptureToggleBusy;
        private string? _lastCaptureStatus;
        private readonly object _hostLogLock = new();
        private readonly string _hostLogFilePath;
        private string? _workerLogFilePath;
        private Button? _selectedTeamButton;
        private Process? _liveTranscriberProcess;

        public MainWindow()
        {
            InitializeComponent();
            _hostLogFilePath = BuildSessionLogFilePath("host");
            LogHost($"MainWindow initialized. Host log path: {_hostLogFilePath}");

            ExtendsContentIntoTitleBar = true;
            SetTitleBar(TitleBarDragRegion);

            AppWindowTitleBar titleBar = AppWindow.TitleBar;
            titleBar.ButtonBackgroundColor = Colors.Transparent;
            titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;

            RootGrid.Loaded += OnRootGridLoaded;
            SizeChanged += OnWindowSizeChanged;
            SidebarCollapseButton.Click += OnSidebarCollapseButtonClick;
            MobileMenuButton.Click += OnMobileMenuButtonClick;
            AddTeamButton.Click += OnAddTeamButtonClick;
            HomeNavButton.Click += OnHomeNavButtonClick;
            QuickNoteButton.Click += OnQuickNoteButtonClick;
            NotesBackButton.Click += OnNotesBackButtonClick;
            SettingsBackButton.Click += OnSettingsBackButtonClick;
            SettingsMenuItem.Click += OnSettingsMenuItemClick;
            HelpCenterMenuItem.Click += OnHelpCenterMenuItemClick;
            CaptureToggleButton.Click += OnCaptureToggleButtonClick;
            Closed += OnMainWindowClosed;

            foreach (Button button in TeamListPanel.Children.OfType<Button>())
            {
                button.Click += OnTeamButtonClick;
                SetTeamButtonState(button, isSelected: false);
            }
        }

        private void OnRootGridLoaded(object sender, RoutedEventArgs e)
        {
            ShowHomePage();
            UpdateLayoutForWidth(AppWindow.Size.Width);
        }

        private void OnWindowSizeChanged(object sender, WindowSizeChangedEventArgs args)
        {
            UpdateLayoutForWidth(args.Size.Width);
        }

        private void UpdateLayoutForWidth(double width)
        {
            bool isWide = width >= WideLayoutBreakpoint;
            bool showSidebar = !_isInNotesMode && !_isInSettingsMode && isWide && !_isSidebarCollapsedByUser;

            SidebarColumn.Width = showSidebar ? new GridLength(210) : new GridLength(0);
            SidebarRoot.Visibility = showSidebar ? Visibility.Visible : Visibility.Collapsed;
            MobileMenuButton.Visibility = showSidebar || _isInNotesMode || _isInSettingsMode ? Visibility.Collapsed : Visibility.Visible;
            QuickNoteButton.Visibility = _isInNotesMode || _isInSettingsMode ? Visibility.Collapsed : Visibility.Visible;
            NotesTitleActions.Visibility = _isInNotesMode ? Visibility.Visible : Visibility.Collapsed;
            SettingsTopTitle.Visibility = _isInSettingsMode ? Visibility.Visible : Visibility.Collapsed;

            if (_isInSettingsMode)
            {
                MainPane.Margin = new Thickness(0, 0, 0, 0);
            }
            else
            {
                MainPane.Margin = showSidebar
                    ? new Thickness(44, 8, 44, 16)
                    : new Thickness(24, 8, 24, 16);
            }

            ContentStack.HorizontalAlignment = isWide
                ? HorizontalAlignment.Center
                : HorizontalAlignment.Stretch;

            ContentStack.MaxWidth = isWide ? 740 : 10000;

            QuickNoteButton.Margin = showSidebar
                ? new Thickness(0, 0, 164, 0)
                : new Thickness(0, 0, 152, 0);
        }

        private void OnSidebarCollapseButtonClick(object sender, RoutedEventArgs e)
        {
            _isSidebarCollapsedByUser = true;
            UpdateLayoutForWidth(AppWindow.Size.Width);
        }

        private void OnMobileMenuButtonClick(object sender, RoutedEventArgs e)
        {
            if (AppWindow.Size.Width >= WideLayoutBreakpoint)
            {
                _isSidebarCollapsedByUser = false;
            }

            UpdateLayoutForWidth(AppWindow.Size.Width);
        }

        private void OnHomeNavButtonClick(object sender, RoutedEventArgs e)
        {
            ShowHomePage();
        }

        private void OnQuickNoteButtonClick(object sender, RoutedEventArgs e)
        {
            ShowNotesPage();
        }

        private async void OnNotesBackButtonClick(object sender, RoutedEventArgs e)
        {
            if (_liveTranscriberProcess is not null)
            {
                await StopLiveTranscriptionAsync();
            }

            ShowHomePage();
        }

        private void OnSettingsMenuItemClick(object sender, RoutedEventArgs e)
        {
            ShowSettingsPage();
        }

        private void OnSettingsBackButtonClick(object sender, RoutedEventArgs e)
        {
            ShowHomePage();
        }

        private async void OnCaptureToggleButtonClick(object sender, RoutedEventArgs e)
        {
            if (_isCaptureToggleBusy)
            {
                return;
            }

            _isCaptureToggleBusy = true;
            CaptureToggleButton.IsEnabled = false;

            try
            {
                if (_liveTranscriberProcess is null)
                {
                    await StartLiveTranscriptionAsync();
                }
                else
                {
                    await StopLiveTranscriptionAsync();
                }
            }
            finally
            {
                _isCaptureToggleBusy = false;
                CaptureToggleButton.IsEnabled = true;
            }
        }

        private async void OnMainWindowClosed(object sender, WindowEventArgs args)
        {
            await StopLiveTranscriptionAsync();
        }

        private async void OnHelpCenterMenuItemClick(object sender, RoutedEventArgs e)
        {
            ContentDialog dialog = new()
            {
                XamlRoot = RootGrid.XamlRoot,
                Title = "Help Center",
                Content = "Help Center placeholder.",
                CloseButtonText = "Close",
            };

            await dialog.ShowAsync();
        }

        private async void OnAddTeamButtonClick(object sender, RoutedEventArgs e)
        {
            TextBox nameBox = new()
            {
                PlaceholderText = "Team name",
                MinWidth = 280,
            };

            ContentDialog dialog = new()
            {
                XamlRoot = RootGrid.XamlRoot,
                Title = "Create team",
                Content = nameBox,
                PrimaryButtonText = "Create",
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Primary,
            };

            ContentDialogResult result = await dialog.ShowAsync();
            if (result != ContentDialogResult.Primary)
            {
                return;
            }

            string teamName = nameBox.Text.Trim();
            if (teamName.Length == 0)
            {
                teamName = $"New team {_nextCreatedTeamNumber++}";
            }

            Button teamButton = CreateTeamButton(teamName);
            TeamListPanel.Children.Add(teamButton);

            SelectTeam(teamButton);
            ShowTeamPage(teamName);
        }

        private void OnTeamButtonClick(object sender, RoutedEventArgs e)
        {
            if (sender is not Button teamButton)
            {
                return;
            }

            string teamName = teamButton.Tag as string ?? "Team";
            SelectTeam(teamButton);
            ShowTeamPage(teamName);
        }

        private Button CreateTeamButton(string teamName)
        {
            Windows.UI.Color teamColor = _teamAccentColors[_nextTeamColorIndex % _teamAccentColors.Length];
            _nextTeamColorIndex++;

            StackPanel row = new()
            {
                Orientation = Orientation.Horizontal,
                Spacing = 6,
            };
            row.Children.Add(new Ellipse
            {
                Width = 5,
                Height = 5,
                VerticalAlignment = VerticalAlignment.Center,
                Fill = new SolidColorBrush(teamColor),
            });
            row.Children.Add(new TextBlock
            {
                Text = teamName,
                FontSize = 12,
                Foreground = new SolidColorBrush(ColorHelper.FromArgb(0xFF, 0x89, 0x90, 0x9A)),
            });

            Button button = new()
            {
                Tag = teamName,
                Content = row,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(6, 2, 6, 2),
                Background = new SolidColorBrush(Colors.Transparent),
                BorderThickness = new Thickness(0),
                CornerRadius = new CornerRadius(6),
            };

            button.Click += OnTeamButtonClick;
            return button;
        }

        private void ShowHomePage()
        {
            _isInNotesMode = false;
            _isInSettingsMode = false;
            HomePageView.Visibility = Visibility.Visible;
            TeamPageView.Visibility = Visibility.Collapsed;
            NotesPageView.Visibility = Visibility.Collapsed;
            SettingsPageView.Visibility = Visibility.Collapsed;
            DeselectCurrentTeam();
            UpdateLayoutForWidth(AppWindow.Size.Width);
        }

        private void ShowTeamPage(string teamName)
        {
            _isInNotesMode = false;
            _isInSettingsMode = false;
            HomePageView.Visibility = Visibility.Collapsed;
            TeamPageView.Visibility = Visibility.Visible;
            NotesPageView.Visibility = Visibility.Collapsed;
            SettingsPageView.Visibility = Visibility.Collapsed;
            TeamTitleText.Text = teamName;
            TeamFolderScopeText.Text = teamName;
            TeamMeetingTitleText.Text = $"{teamName} / ProSourceIT AI Weekly Standup (Teams Meeting Only)";
            UpdateLayoutForWidth(AppWindow.Size.Width);
        }

        private void ShowNotesPage()
        {
            _isInNotesMode = true;
            _isInSettingsMode = false;
            HomePageView.Visibility = Visibility.Collapsed;
            TeamPageView.Visibility = Visibility.Collapsed;
            NotesPageView.Visibility = Visibility.Visible;
            SettingsPageView.Visibility = Visibility.Collapsed;
            DeselectCurrentTeam();
            UpdateLayoutForWidth(AppWindow.Size.Width);
            NoteTitleTextBox.Focus(FocusState.Programmatic);
        }

        private void ShowSettingsPage()
        {
            _isInNotesMode = false;
            _isInSettingsMode = true;
            HomePageView.Visibility = Visibility.Collapsed;
            TeamPageView.Visibility = Visibility.Collapsed;
            NotesPageView.Visibility = Visibility.Collapsed;
            SettingsPageView.Visibility = Visibility.Visible;
            DeselectCurrentTeam();
            UpdateLayoutForWidth(AppWindow.Size.Width);
        }

        private void SelectTeam(Button selectedButton)
        {
            if (_selectedTeamButton == selectedButton)
            {
                return;
            }

            if (_selectedTeamButton is not null)
            {
                SetTeamButtonState(_selectedTeamButton, isSelected: false);
            }

            _selectedTeamButton = selectedButton;
            SetTeamButtonState(selectedButton, isSelected: true);
        }

        private void DeselectCurrentTeam()
        {
            if (_selectedTeamButton is not null)
            {
                SetTeamButtonState(_selectedTeamButton, isSelected: false);
                _selectedTeamButton = null;
            }
        }

        private static void SetTeamButtonState(Button button, bool isSelected)
        {
            button.Background = isSelected
                ? new SolidColorBrush(ColorHelper.FromArgb(0xFF, 0x2B, 0x2E, 0x33))
                : new SolidColorBrush(Colors.Transparent);

            if (button.Content is StackPanel row && row.Children.Count > 1 && row.Children[1] is TextBlock label)
            {
                label.Foreground = isSelected
                    ? new SolidColorBrush(ColorHelper.FromArgb(0xFF, 0xB7, 0xD1, 0xEA))
                    : new SolidColorBrush(ColorHelper.FromArgb(0xFF, 0x89, 0x90, 0x9A));
            }
        }

        private async Task StartLiveTranscriptionAsync()
        {
            _workerLogFilePath = BuildSessionLogFilePath("worker");
            string? rustWorkerPath = ResolveRustWorkerPath();
            string ggmlModelPath = ResolveGgmlModelPath();
            bool canUseRust = rustWorkerPath is not null && File.Exists(ggmlModelPath);
            bool disableMic = ResolveBooleanEnvironmentVariable("WHISPER_DISABLE_MIC");
            bool disableLoopback = ResolveBooleanEnvironmentVariable("WHISPER_DISABLE_LOOPBACK");
            string sourceArgs = $"{(disableMic ? " --disable-mic" : string.Empty)}{(disableLoopback ? " --disable-loopback" : string.Empty)}";

            ProcessStartInfo startInfo;
            string backendName;

            if (canUseRust)
            {
                string selectedRustWorkerPath = rustWorkerPath!;
                backendName = "rust";
                startInfo = new ProcessStartInfo
                {
                    FileName = selectedRustWorkerPath,
                    Arguments = $"--model-path \"{ggmlModelPath}\" --device cuda --chunk-seconds 6 --log-file \"{_workerLogFilePath}\" --verbose-chunk-log{sourceArgs}",
                    WorkingDirectory = IOPath.GetDirectoryName(selectedRustWorkerPath) ?? AppContext.BaseDirectory,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                };

                LogHost($"Start requested. Using rust worker: {selectedRustWorkerPath}. GGML model path: {ggmlModelPath}");
            }
            else
            {
                string scriptPath = ResolveTranscriberScriptPath();
                if (rustWorkerPath is null)
                {
                    LogHost($"Start requested. Rust worker not found; falling back to python script: {scriptPath}");
                }
                else
                {
                    LogHost(
                        $"Start requested. Rust worker found at {rustWorkerPath}, but GGML model file was not found at {ggmlModelPath}. Falling back to python script: {scriptPath}");
                }

                if (!File.Exists(scriptPath))
                {
                    SetCaptureStatus("No transcriber backend found. Build rust worker or verify scripts/live_meeting_transcriber.py.");
                    LogHost("Start aborted. Python transcriber script file not found.");
                    return;
                }

                string pythonExecutable = Environment.GetEnvironmentVariable("WHISPER_PYTHON") ?? "python";
                string whisperModel = ResolveWhisperModel();

                backendName = "python";
                startInfo = new ProcessStartInfo
                {
                    FileName = pythonExecutable,
                    Arguments = $"-u \"{scriptPath}\" --model {whisperModel} --device cuda --chunk-seconds 6 --log-file \"{_workerLogFilePath}\" --verbose-chunk-log{sourceArgs}",
                    WorkingDirectory = IOPath.GetDirectoryName(scriptPath) ?? AppContext.BaseDirectory,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                };

                startInfo.Environment["PYTHONUNBUFFERED"] = "1";
                startInfo.Environment["HF_HUB_DISABLE_PROGRESS_BARS"] = "1";
                LogHost($"Python fallback selected. Python: {pythonExecutable}. Model: {whisperModel}. Worker log: {_workerLogFilePath}");
            }

            LogHost($"Capture source configuration. disableMic={disableMic}, disableLoopback={disableLoopback}");

            Process process = new()
            {
                StartInfo = startInfo,
                EnableRaisingEvents = true,
            };

            process.OutputDataReceived += OnLiveTranscriberOutputDataReceived;
            process.ErrorDataReceived += OnLiveTranscriberErrorDataReceived;
            process.Exited += OnLiveTranscriberExited;
            LogHost($"Launching transcriber backend={backendName}. Executable: {process.StartInfo.FileName}. Worker log: {_workerLogFilePath}");

            try
            {
                bool started = process.Start();
                if (!started)
                {
                    SetCaptureStatus("Unable to start transcription process.");
                    LogHost("Start failed. process.Start returned false.");
                    return;
                }
            }
            catch (Exception ex)
            {
                SetCaptureStatus($"Failed to start transcriber: {ex.Message}");
                LogHost($"Start failed with exception: {ex}");
                return;
            }

            _liveTranscriberProcess = process;
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            UpdateCaptureUi(isCapturing: true);
            SetCaptureStatus("Starting live transcript...");
            LogHost($"Transcriber started. PID: {process.Id}");

            await Task.CompletedTask;
        }

        private async Task StopLiveTranscriptionAsync()
        {
            Process? process = _liveTranscriberProcess;
            if (process is null)
            {
                UpdateCaptureUi(isCapturing: false);
                LogHost("Stop requested with no active transcriber process.");
                return;
            }

            _liveTranscriberProcess = null;
            UpdateCaptureUi(isCapturing: false);
            SetCaptureStatus("Stopping live transcript...");
            LogHost($"Stop requested. PID: {process.Id}");

            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                    await process.WaitForExitAsync();
                    LogHost($"Transcriber process killed and awaited. PID: {process.Id}");
                }
            }
            catch (Exception ex)
            {
                SetCaptureStatus($"Stop failed: {ex.Message}");
                LogHost($"Stop failed with exception: {ex}");
            }
            finally
            {
                process.OutputDataReceived -= OnLiveTranscriberOutputDataReceived;
                process.ErrorDataReceived -= OnLiveTranscriberErrorDataReceived;
                process.Exited -= OnLiveTranscriberExited;
                process.Dispose();
                LogHost("Transcriber process resources disposed.");
            }

            SetCaptureStatus("Live transcript stopped.");
            LogHost("Stop completed.");
        }

        private void OnLiveTranscriberOutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Data))
            {
                return;
            }

            string line = e.Data.Trim();
            LogHost($"Worker stdout: {line}");
            HandleTranscriberMessage(line);
        }

        private void OnLiveTranscriberErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Data))
            {
                return;
            }

            string line = e.Data.Trim();
            LogHost($"Worker stderr: {line}");
            if (line.Contains("error", StringComparison.OrdinalIgnoreCase)
                || line.Contains("exception", StringComparison.OrdinalIgnoreCase)
                || line.Contains("traceback", StringComparison.OrdinalIgnoreCase))
            {
                SetCaptureStatus(line);
            }
        }

        private void OnLiveTranscriberExited(object? sender, EventArgs e)
        {
            if (sender is not Process process)
            {
                return;
            }

            if (!ReferenceEquals(_liveTranscriberProcess, process))
            {
                LogHost($"Received exit event from stale process. PID: {process.Id}");
                return;
            }

            _liveTranscriberProcess = null;
            UpdateCaptureUi(isCapturing: false);

            int exitCode;
            try
            {
                exitCode = process.ExitCode;
            }
            catch
            {
                SetCaptureStatus("Live transcript process exited.");
                LogHost($"Transcriber process exit observed without exit code. PID: {process.Id}");
                return;
            }

            SetCaptureStatus(exitCode == 0
                ? "Live transcript stopped."
                : $"Live transcript exited with code {exitCode}.");
            LogHost($"Transcriber exited. PID: {process.Id}, ExitCode: {exitCode}");
        }

        private void HandleTranscriberMessage(string line)
        {
            try
            {
                using JsonDocument document = JsonDocument.Parse(line);
                JsonElement root = document.RootElement;

                if (!root.TryGetProperty("event", out JsonElement eventNameElement))
                {
                    return;
                }

                string eventName = eventNameElement.GetString() ?? string.Empty;
                switch (eventName)
                {
                    case "status":
                        if (root.TryGetProperty("message", out JsonElement statusMessageElement))
                        {
                            string status = statusMessageElement.GetString() ?? "Status update";
                            SetCaptureStatus(status);
                            LogHost($"Worker status: {status}");
                        }

                        break;
                    case "transcript":
                        if (root.TryGetProperty("text", out JsonElement textElement))
                        {
                            string transcript = textElement.GetString() ?? string.Empty;
                            LogHost($"Worker transcript text length: {transcript.Length}");
                            AppendTranscriptText(transcript);
                        }

                        break;
                    case "error":
                        if (root.TryGetProperty("message", out JsonElement errorMessageElement))
                        {
                            string error = errorMessageElement.GetString() ?? "Unknown transcriber error.";
                            SetCaptureStatus(error);
                            LogHost($"Worker error event: {error}");
                        }

                        break;
                }
            }
            catch (JsonException)
            {
                LogHost($"Worker stdout non-JSON line: {line}");
                if (line.Contains("error", StringComparison.OrdinalIgnoreCase)
                    || line.Contains("exception", StringComparison.OrdinalIgnoreCase)
                    || line.Contains("traceback", StringComparison.OrdinalIgnoreCase))
                {
                    SetCaptureStatus(line);
                }
            }
        }

        private void AppendTranscriptText(string text)
        {
            string normalized = text.Trim();
            if (normalized.Length == 0)
            {
                return;
            }

            RunOnUiThread(() =>
            {
                if (NoteBodyTextBox.Text.Length > 0 && !NoteBodyTextBox.Text.EndsWith(Environment.NewLine, StringComparison.Ordinal))
                {
                    NoteBodyTextBox.Text += Environment.NewLine;
                }

                NoteBodyTextBox.Text += normalized + Environment.NewLine;
            });
            LogHost($"Transcript appended. Characters: {normalized.Length}");
        }

        private void SetCaptureStatus(string status)
        {
            string normalized = status.Trim();
            if (normalized.Length == 0)
            {
                return;
            }

            if (normalized.Length > 220)
            {
                normalized = normalized[..220];
            }

            if (string.Equals(_lastCaptureStatus, normalized, StringComparison.Ordinal))
            {
                return;
            }

            _lastCaptureStatus = normalized;
            RunOnUiThread(() => CaptureStatusText.Text = normalized);
            LogHost($"Capture status changed: {normalized}");
        }

        private void UpdateCaptureUi(bool isCapturing)
        {
            RunOnUiThread(() =>
            {
                CaptureButtonText.Text = isCapturing ? "Stop live transcript" : "Start live transcript";
                CaptureIndicatorText.Foreground = new SolidColorBrush(isCapturing
                    ? ColorHelper.FromArgb(0xFF, 0xFF, 0x72, 0x72)
                    : ColorHelper.FromArgb(0xFF, 0x9B, 0xCB, 0x6D));
            });
        }

        private void RunOnUiThread(Action callback)
        {
            _ = DispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    callback();
                }
                catch (Exception ex) when (IsBenignUiInteropException(ex))
                {
                    Debug.WriteLine($"Ignored benign UI interop exception: 0x{ex.HResult:X8}");
                    LogHost($"Ignored benign UI interop exception: 0x{ex.HResult:X8}");
                }
            });
        }

        private static bool IsBenignUiInteropException(Exception ex)
        {
            if (ex is not COMException comException)
            {
                return false;
            }

            return comException.HResult == unchecked((int)0x8001010D) // RPC_E_CANTCALLOUT_ININPUTSYNCCALL
                || comException.HResult == unchecked((int)0x80070057)  // E_INVALIDARG
                || comException.HResult == unchecked((int)0x80004005); // E_FAIL
        }

        private static string ResolveTranscriberScriptPath()
        {
            string bundledPath = IOPath.Combine(AppContext.BaseDirectory, "scripts", "live_meeting_transcriber.py");
            if (File.Exists(bundledPath))
            {
                return bundledPath;
            }

            return IOPath.GetFullPath(IOPath.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "scripts", "live_meeting_transcriber.py"));
        }

        private static string ResolveWhisperModel()
        {
            string? overrideModel = Environment.GetEnvironmentVariable("WHISPER_MODEL");
            if (!string.IsNullOrWhiteSpace(overrideModel))
            {
                return overrideModel.Trim();
            }

            return DefaultWhisperModel;
        }

        private static bool ResolveBooleanEnvironmentVariable(string variableName)
        {
            string? value = Environment.GetEnvironmentVariable(variableName);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.Trim().ToLowerInvariant() switch
            {
                "1" => true,
                "true" => true,
                "yes" => true,
                "on" => true,
                _ => false,
            };
        }

        private static string ResolveGgmlModelPath()
        {
            string? overridePath = Environment.GetEnvironmentVariable("WHISPER_GGML_MODEL_PATH");
            if (!string.IsNullOrWhiteSpace(overridePath))
            {
                return overridePath.Trim();
            }

            string bundledPath = IOPath.Combine(AppContext.BaseDirectory, "models", "ggml-medium.bin");
            if (File.Exists(bundledPath))
            {
                return bundledPath;
            }

            return IOPath.GetFullPath(IOPath.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "models", "ggml-medium.bin"));
        }

        private static string? ResolveRustWorkerPath()
        {
            string? overridePath = Environment.GetEnvironmentVariable("WHISPER_RUST_WORKER");
            if (!string.IsNullOrWhiteSpace(overridePath) && File.Exists(overridePath))
            {
                return overridePath.Trim();
            }

            string bundledPath = IOPath.Combine(AppContext.BaseDirectory, "rust_worker", "live_meeting_transcriber_rust.exe");
            if (File.Exists(bundledPath))
            {
                return bundledPath;
            }

            string releasePath = IOPath.GetFullPath(IOPath.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                "rust-worker",
                "target",
                "release",
                "live_meeting_transcriber_rust.exe"));
            if (File.Exists(releasePath))
            {
                return releasePath;
            }

            string debugPath = IOPath.GetFullPath(IOPath.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                "rust-worker",
                "target",
                "debug",
                "live_meeting_transcriber_rust.exe"));
            if (File.Exists(debugPath))
            {
                return debugPath;
            }

            return null;
        }

        private static string BuildSessionLogFilePath(string prefix)
        {
            string directory = IOPath.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Whisper",
                "logs");
            Directory.CreateDirectory(directory);

            string timestamp = DateTimeOffset.Now.ToString("yyyyMMdd-HHmmss");
            string fileName = $"{prefix}-{timestamp}-{Environment.ProcessId}.log";
            return IOPath.Combine(directory, fileName);
        }

        private void LogHost(string message)
        {
            string line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff} [host] [tid:{Environment.CurrentManagedThreadId}] {message}";
            Debug.WriteLine(line);

            try
            {
                lock (_hostLogLock)
                {
                    File.AppendAllText(_hostLogFilePath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never crash transcription flow.
            }
        }
    }
}
