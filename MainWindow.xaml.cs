using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms; // NotifyIcon
using System.Windows.Input;
using System.Windows.Media;
using BRDesktopAssistant.Models;
using BRDesktopAssistant.Services;
using BRDesktopAssistant.Utils;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace BRDesktopAssistant
{
    public partial class MainWindow : Window
    {
        private readonly OpenAIClient _client;
        private readonly ChatMessage[] _system;
        private readonly SpeechSynthesizer _tts = new SpeechSynthesizer();
        private SpeechRecognitionEngine? _stt;
        private readonly NotifyIcon _tray;
        private bool _startWithWindows = false;
        private readonly string _appData;
        private readonly string _docsDir;

        public MainWindow()
        {
            InitializeComponent();

            _appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BRDesktopAssistant");
            Directory.CreateDirectory(_appData);
            _docsDir = Path.Combine(_appData, "Documents");
            Directory.CreateDirectory(_docsDir);

            // OpenAI
            _client = new OpenAIClient();
            _system = new[]
            {
                new ChatMessage("system", "Ты — голосовой/текстовый ассистент Б&Р (Бизнес и Разум). Помогаешь предпринимателю: создаёшь счета, накладные, акты (DOCX), даёшь советы по автоматизации и продажам. Отвечай кратко, шагами, на русском.")
            };

            // TTS
            try
            {
                _tts.Rate = 0;
                _tts.Volume = 100;
                // Попытка выбрать русскую голосовую
                var ru = _tts.GetInstalledVoices().FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("ru", StringComparison.OrdinalIgnoreCase));
                if (ru != null) _tts.SelectVoice(ru.VoiceInfo.Name);
            }
            catch { }

            // STT подготовка (не включаем до клика)
            InitTray();
            RegisterHotkey();

            PromptBox.KeyDown += async (s, e) =>
            {
                if (e.Key == Key.Enter && !Keyboard.IsKeyDown(Key.LeftShift))
                {
                    e.Handled = true;
                    await SendAsync();
                }
            };
        }

        private void InitTray()
        {
            _tray = new NotifyIcon();
            _tray.Icon = new System.Drawing.Icon(AppDomain.CurrentDomain.BaseDirectory + "Assets\\br.ico");
            _tray.Visible = true;
            _tray.Text = "Б&Р Ассистент";

            var menu = new ContextMenuStrip();
            menu.Items.Add("Показать окно", null, (_, __) => { this.Show(); this.WindowState = WindowState.Normal; this.Activate(); });
            var ttsItem = new ToolStripMenuItem("Озвучка (TTS)") { Checked = true, CheckOnClick = true };
            ttsItem.CheckedChanged += (_, __) => { TtsToggle.IsChecked = ttsItem.Checked; };
            menu.Items.Add(ttsItem);
            var sttItem = new ToolStripMenuItem("Распознавание (STT)") { Checked = false, CheckOnClick = true };
            sttItem.CheckedChanged += (_, __) => {
                if (sttItem.Checked) StartStt();
                else StopStt();
            };
            menu.Items.Add(sttItem);
            menu.Items.Add(new ToolStripSeparator());
            var autoStart = new ToolStripMenuItem("Автозапуск Windows") { Checked = _startWithWindows, CheckOnClick = true };
            autoStart.CheckedChanged += (_, __) => { _startWithWindows = autoStart.Checked; StartupHelper.SetAutorun(_startWithWindows); };
            menu.Items.Add(autoStart);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Выход", null, (_, __) => { _tray.Visible = false; Application.Current.Shutdown(); });

            _tray.ContextMenuStrip = menu;
            _tray.DoubleClick += (_, __) => { this.Show(); this.WindowState = WindowState.Normal; this.Activate(); };
        }

        private void RegisterHotkey()
        {
            // Ctrl+Alt+B
            HotkeyManager.Register(this, HotkeyModifiers.MOD_CONTROL | HotkeyModifiers.MOD_ALT, System.Windows.Forms.Keys.B, () =>
            {
                this.Show();
                this.WindowState = WindowState.Normal;
                this.Activate();
                PromptBox.Focus();
            });
        }

        private Border MakeBubble(string text, bool isUser)
        {
            var tb = new TextBlock
            {
                Text = text,
                TextWrapping = TextWrapping.Wrap,
                Foreground = Brushes.White,
                Margin = new Thickness(10)
            };

            return new Border
            {
                Background = (isUser ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1f2937"))
                                      : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827"))),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(4),
                Margin = new Thickness(0, 6, 0, 6),
                Child = tb,
                HorizontalAlignment = isUser ? HorizontalAlignment.Right : HorizontalAlignment.Left,
                MaxWidth = 760
            };
        }

        private async Task SendAsync()
        {
            var prompt = PromptBox.Text?.Trim();
            if (string.IsNullOrEmpty(prompt)) return;

            HistoryPanel.Children.Add(MakeBubble(prompt, true));
            PromptBox.Clear();
            HistoryScroll.ScrollToEnd();

            try
            {
                var messages = _system.Concat(new[] { new ChatMessage("user", prompt) }).ToArray();
                var reply = await _client.GetChatCompletionAsync(messages);

                HistoryPanel.Children.Add(MakeBubble(reply, false));
                HistoryScroll.ScrollToEnd();

                if (TtsToggle.IsChecked == true)
                {
                    try { _tts.SpeakAsyncCancelAll(); _tts.SpeakAsync(reply); } catch { }
                }
            }
            catch (Exception ex)
            {
                HistoryPanel.Children.Add(MakeBubble($"Ошибка: {ex.Message}", false));
                HistoryScroll.ScrollToEnd();
            }
        }

        // ==== STT ====
        private void StartStt()
        {
            try
            {
                if (_stt != null) return;
                var culture = new CultureInfo("ru-RU");
                try { _stt = new SpeechRecognitionEngine(culture); }
                catch { _stt = new SpeechRecognitionEngine(new CultureInfo("en-US")); }

                _stt.SetInputToDefaultAudioDevice();
                var dict = new Choices("создай счёт", "создай накладную", "создай акт", "отправить", "позвать бир");
                var gb = new GrammarBuilder(dict) { Culture = _stt.RecognizerInfo.Culture };
                var grammar = new Grammar(gb);
                _stt.LoadGrammar(grammar);
                _stt.SpeechRecognized += (s, e) =>
                {
                    var text = e.Result?.Text ?? "";
                    if (text.Contains("счёт")) Invoice_Click(this, new RoutedEventArgs());
                    else if (text.Contains("накладную")) Waybill_Click(this, new RoutedEventArgs());
                    else if (text.Contains("акт")) Act_Click(this, new RoutedEventArgs());
                    else if (text.Contains("отправить")) _ = SendAsync();
                    else if (text.Contains("бир")) HotkeyManager.Trigger();
                };
                _stt.RecognizeAsync(RecognizeMode.Multiple);
                SttToggle.IsChecked = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("STT не удалось запустить: " + ex.Message);
                SttToggle.IsChecked = false;
            }
        }

        private void StopStt()
        {
            try
            {
                _stt?.RecognizeAsyncStop();
                _stt = null;
            }
            catch { }
        }

        private void SttToggle_Checked(object sender, RoutedEventArgs e) => StartStt();
        private void SttToggle_Unchecked(object sender, RoutedEventArgs e) => StopStt();

        // ==== Документы ====
        private void Invoice_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var path = DocService.CreateInvoice(_docsDir, "Орлов", "Ерохин", "Подписка на ассистента", 1, 30000m);
                HistoryPanel.Children.Add(MakeBubble($"Счёт создан: {path}", false));
                HistoryScroll.ScrollToEnd();
            }
            catch (Exception ex) { HistoryPanel.Children.Add(MakeBubble($"Ошибка счёта: {ex.Message}", false)); }
        }
        private void Waybill_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var path = DocService.CreateWaybill(_docsDir, "Орлов", "Ерохин", "Подписка на ассистента", 1, 30000m);
                HistoryPanel.Children.Add(MakeBubble($"Накладная создана: {path}", false));
                HistoryScroll.ScrollToEnd();
            }
            catch (Exception ex) { HistoryPanel.Children.Add(MakeBubble($"Ошибка накладной: {ex.Message}", false)); }
        }
        private void Act_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var path = DocService.CreateAct(_docsDir, "Орлов Кирилл Евгеньевич", "Ерохин Андрей Николаевич", "Подписка на ассистента", 30000m);
                HistoryPanel.Children.Add(MakeBubble($"Акт создан: {path}", false));
                HistoryScroll.ScrollToEnd();
            }
            catch (Exception ex) { HistoryPanel.Children.Add(MakeBubble($"Ошибка акта: {ex.Message}", false)); }
        }

        protected override void OnStateChanged(EventArgs e)
        {
            base.OnStateChanged(e);
            if (WindowState == WindowState.Minimized) this.Hide();
        }

        protected override void OnClosed(EventArgs e)
        {
            _tray.Visible = false;
            HotkeyManager.Unregister(this);
            base.OnClosed(e);
        }
    }
}
