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
using System.Text.Json;
using System.Threading.Tasks;

namespace Whisper
{
    public sealed partial class MainWindow : Window
    {
        private const double WideLayoutBreakpoint = 980;
        private const string DefaultWhisperModel = "openai/whisper-large-v3-turbo";
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
        private Button? _selectedTeamButton;
        private Process? _liveTranscriberProcess;

        public MainWindow()
        {
            InitializeComponent();

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
            string scriptPath = ResolveTranscriberScriptPath();
            if (!File.Exists(scriptPath))
            {
                SetCaptureStatus("Transcriber script not found. Build once or check scripts/live_meeting_transcriber.py.");
                return;
            }

            string pythonExecutable = Environment.GetEnvironmentVariable("WHISPER_PYTHON") ?? "python";
            Process process = new()
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = pythonExecutable,
                    Arguments = $"\"{scriptPath}\" --model {DefaultWhisperModel} --device cuda --chunk-seconds 6",
                    WorkingDirectory = Path.GetDirectoryName(scriptPath) ?? AppContext.BaseDirectory,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                },
                EnableRaisingEvents = true,
            };

            process.OutputDataReceived += OnLiveTranscriberOutputDataReceived;
            process.ErrorDataReceived += OnLiveTranscriberErrorDataReceived;
            process.Exited += OnLiveTranscriberExited;

            try
            {
                bool started = process.Start();
                if (!started)
                {
                    SetCaptureStatus("Unable to start transcription process.");
                    return;
                }
            }
            catch (Exception ex)
            {
                SetCaptureStatus($"Failed to start transcriber: {ex.Message}");
                return;
            }

            _liveTranscriberProcess = process;
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            UpdateCaptureUi(isCapturing: true);
            SetCaptureStatus("Starting live transcript...");

            await Task.CompletedTask;
        }

        private async Task StopLiveTranscriptionAsync()
        {
            Process? process = _liveTranscriberProcess;
            if (process is null)
            {
                UpdateCaptureUi(isCapturing: false);
                return;
            }

            _liveTranscriberProcess = null;
            UpdateCaptureUi(isCapturing: false);
            SetCaptureStatus("Stopping live transcript...");

            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                    await process.WaitForExitAsync();
                }
            }
            catch (Exception ex)
            {
                SetCaptureStatus($"Stop failed: {ex.Message}");
            }
            finally
            {
                process.OutputDataReceived -= OnLiveTranscriberOutputDataReceived;
                process.ErrorDataReceived -= OnLiveTranscriberErrorDataReceived;
                process.Exited -= OnLiveTranscriberExited;
                process.Dispose();
            }

            SetCaptureStatus("Live transcript stopped.");
        }

        private void OnLiveTranscriberOutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Data))
            {
                return;
            }

            HandleTranscriberMessage(e.Data);
        }

        private void OnLiveTranscriberErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Data))
            {
                return;
            }

            SetCaptureStatus(e.Data);
        }

        private void OnLiveTranscriberExited(object? sender, EventArgs e)
        {
            if (sender is not Process process)
            {
                return;
            }

            if (_liveTranscriberProcess is not process)
            {
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
                return;
            }

            SetCaptureStatus(exitCode == 0
                ? "Live transcript stopped."
                : $"Live transcript exited with code {exitCode}.");
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
                            SetCaptureStatus(statusMessageElement.GetString() ?? "Status update");
                        }

                        break;
                    case "transcript":
                        if (root.TryGetProperty("text", out JsonElement textElement))
                        {
                            AppendTranscriptText(textElement.GetString() ?? string.Empty);
                        }

                        break;
                    case "error":
                        if (root.TryGetProperty("message", out JsonElement errorMessageElement))
                        {
                            SetCaptureStatus(errorMessageElement.GetString() ?? "Unknown transcriber error.");
                        }

                        break;
                }
            }
            catch (JsonException)
            {
                SetCaptureStatus(line);
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
                NoteBodyTextBox.SelectionStart = NoteBodyTextBox.Text.Length;
                NoteBodyTextBox.SelectionLength = 0;
            });
        }

        private void SetCaptureStatus(string status)
        {
            RunOnUiThread(() => CaptureStatusText.Text = status);
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
            if (DispatcherQueue.HasThreadAccess)
            {
                callback();
            }
            else
            {
                _ = DispatcherQueue.TryEnqueue(() => callback());
            }
        }

        private static string ResolveTranscriberScriptPath()
        {
            string bundledPath = Path.Combine(AppContext.BaseDirectory, "scripts", "live_meeting_transcriber.py");
            if (File.Exists(bundledPath))
            {
                return bundledPath;
            }

            return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "scripts", "live_meeting_transcriber.py"));
        }
    }
}
