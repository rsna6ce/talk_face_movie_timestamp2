using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NAudio.Wave;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

// BudouX (Apache License 2.0) - Parser.csをプロジェクトに追加済みとして使用
using BudouX;

namespace talk_face_movie_timestamp2
{
    public partial class Form1 : Form
    {
        private TextBox txtInputFolder;
        private TextBox txtOutputWav;
        private TextBox txtOutputCsv;
        private TextBox txtVoiceActor;
        private TextBox txtResult;
        private Button btnSelectInputFolder;
        private Button btnSelectOutputWav;
        private Button btnSelectOutputCsv;
        private Button btnStart;
        private Button btnCleanup;
        private CheckBox chkAutoCleanup;
        private CheckBox chkAutoAss; // ← 新規追加

        private bool autoSuccess = false;

        private readonly string settingsFilePath = Path.Combine(Application.StartupPath, "settings.json");

        public Form1()
        {
            InitializeComponent();
            InitializeUI();
        }

        private void InitializeUI()
        {
            // 入力フォルダ
            Label lblInputFolder = new Label { Text = "入力フォルダ:", Location = new System.Drawing.Point(10, 10), Width = 100 };
            lblInputFolder.DoubleClick += lblInputFolder_DoubleClick;
            txtInputFolder = new TextBox { Location = new System.Drawing.Point(120, 10), Width = 600, Font = new System.Drawing.Font("ＭＳ ゴシック", 9) };
            btnSelectInputFolder = new Button { Text = "選択", Location = new System.Drawing.Point(730, 10), Width = 60 };
            btnSelectInputFolder.Click += BtnSelectInputFolder_Click;

            // 出力WAVファイル
            Label lblOutputWav = new Label { Text = "出力WAV:", Location = new System.Drawing.Point(10, 40), Width = 100 };
            txtOutputWav = new TextBox { Location = new System.Drawing.Point(120, 40), Width = 600, Font = new System.Drawing.Font("ＭＳ ゴシック", 9) };
            btnSelectOutputWav = new Button { Text = "選択", Location = new System.Drawing.Point(730, 40), Width = 60 };
            btnSelectOutputWav.Click += BtnSelectOutputWav_Click;

            // 出力CSVファイル
            Label lblOutputCsv = new Label { Text = "出力CSV:", Location = new System.Drawing.Point(10, 70), Width = 100 };
            txtOutputCsv = new TextBox { Location = new System.Drawing.Point(120, 70), Width = 600, Font = new System.Drawing.Font("ＭＳ ゴシック", 9) };
            btnSelectOutputCsv = new Button { Text = "選択", Location = new System.Drawing.Point(730, 70), Width = 60 };
            btnSelectOutputCsv.Click += BtnSelectOutputCsv_Click;

            // 声優設定
            Label lblVoiceActor = new Label { Text = "声優設定(★):", Location = new System.Drawing.Point(10, 100), Width = 100 };
            txtVoiceActor = new TextBox { Location = new System.Drawing.Point(120, 100), Width = 600, Font = new System.Drawing.Font("ＭＳ ゴシック", 9), Text = "" };

            // 開始ボタン
            btnStart = new Button { Text = "開始", Location = new System.Drawing.Point(10, 130), Width = 100 };
            btnStart.Click += BtnStart_Click;

            // クリーンアップボタン
            btnCleanup = new Button { Text = "入力フォルダクリーンアップ", Location = new System.Drawing.Point(120, 130), Width = 150 };
            btnCleanup.Click += BtnCleanup_Click;

            // 自動クリーンアップチェックボックス
            chkAutoCleanup = new CheckBox
            {
                Text = "変換後自動クリーンアップ",
                Location = new System.Drawing.Point(280, 130),
                Width = 150,
                Checked = true
            };

            // ← 新規追加：ASS自動生成チェックボックス
            chkAutoAss = new CheckBox
            {
                Text = "ASSファイル自動生成",
                Location = new System.Drawing.Point(440, 130),
                Width = 160,
                Checked = true // default true
            };

            // 結果表示
            Label lblResult = new Label { Text = "結果:", Location = new System.Drawing.Point(10, 160), Width = 100 };
            txtResult = new TextBox { Location = new System.Drawing.Point(10, 190), Width = 780, Height = 350, Multiline = true, ReadOnly = true, Font = new System.Drawing.Font("ＭＳ ゴシック", 9), ScrollBars = ScrollBars.Vertical };

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] { lblInputFolder, txtInputFolder, btnSelectInputFolder,
                                                  lblOutputWav, txtOutputWav, btnSelectOutputWav,
                                                  lblOutputCsv, txtOutputCsv, btnSelectOutputCsv,
                                                  lblVoiceActor, txtVoiceActor, btnStart, btnCleanup,
                                                  chkAutoCleanup, chkAutoAss, // ← 追加
                                                  lblResult, txtResult });

            // イベントハンドラの登録
            this.Load += Form1_Load;
            this.FormClosing += Form1_FormClosing;
        }

        // （Form1_Load, Form1_FormClosing, lblInputFolder_DoubleClick, BtnSelect* は変更なし）
        private void Form1_Load(object sender, EventArgs e)
        {
            // JSONファイルから設定を読み込む
            try
            {
                if (File.Exists(settingsFilePath))
                {
                    using (var stream = new FileStream(settingsFilePath, FileMode.Open, FileAccess.Read))
                    {
                        var serializer = new DataContractJsonSerializer(typeof(Settings));
                        var settings = (Settings)serializer.ReadObject(stream);
                        txtInputFolder.Text = settings.InputFolder ?? "";
                        txtOutputWav.Text = settings.OutputWav ?? "";
                        txtOutputCsv.Text = settings.OutputCsv ?? "";
                        txtVoiceActor.Text = settings.VoiceActor ?? "猫使アル";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"設定の読み込みに失敗: {ex.Message}");
            }

            // ====================== /auto モード処理 ======================
            var args = Environment.GetCommandLineArgs();
            bool autoMode = false;
            string inputPathFromArg = null;

            for (int i = 1; i < args.Length; i++)
            {
                if (args[i].Equals("/auto", StringComparison.OrdinalIgnoreCase))
                {
                    autoMode = true;
                }
                else if (args[i].Equals("/input", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    inputPathFromArg = args[i + 1].Trim('"');
                    i++; // 次の引数をスキップ
                }
            }

            // /input が指定されていたら上書き
            if (!string.IsNullOrEmpty(inputPathFromArg))
            {
                txtInputFolder.Text = inputPathFromArg;

                // 出力ファイル名も自動設定
                txtOutputWav.Text = inputPathFromArg + ".wav";
                txtOutputCsv.Text = inputPathFromArg + ".csv";
            }

            if (autoMode)
            {
                txtResult.Text = "自動実行モード ...";

                // 少し待ってから自動実行（UIが表示されるのを待つ）
                this.Shown += (s, ev) =>
                {
                    System.Threading.Tasks.Task.Delay(800).ContinueWith(t =>
                    {
                        this.Invoke(new Action(() =>
                        {
                            BtnStart_Click(null, null);  // ボタンクリックと同じ処理を実行
                        }));
                    });
                };
            }
            // ============================================================
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 設定をJSONファイルに保存
            try
            {
                var settings = new Settings
                {
                    InputFolder = txtInputFolder.Text,
                    OutputWav = txtOutputWav.Text,
                    OutputCsv = txtOutputCsv.Text,
                    VoiceActor = txtVoiceActor.Text
                };

                using (var stream = new FileStream(settingsFilePath, FileMode.Create, FileAccess.Write))
                {
                    var serializer = new DataContractJsonSerializer(typeof(Settings));
                    serializer.WriteObject(stream, settings);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"設定の保存に失敗: {ex.Message}");
            }

            var args = Environment.GetCommandLineArgs();
            bool isAutoMode = args.Any(arg => arg.Equals("/auto", StringComparison.OrdinalIgnoreCase));
            if (isAutoMode)
            {
                // 成功フラグがfalseならエラー終了コードを返す
                Environment.Exit(autoSuccess ? 0 : 1);
            }
        }

        private void lblInputFolder_DoubleClick(object sender, EventArgs e)
        {
            string directory_name = txtInputFolder.Text;
            if (directory_name != "")
            {
                txtOutputWav.Text = directory_name + ".wav";
                txtOutputCsv.Text = directory_name + ".csv";
            }
        }

        private void BtnSelectInputFolder_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (!string.IsNullOrEmpty(txtInputFolder.Text) && Directory.Exists(txtInputFolder.Text))
                {
                    fbd.SelectedPath = txtInputFolder.Text;
                }

                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtInputFolder.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnSelectOutputWav_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog { Filter = "WAVファイル|*.wav" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtOutputWav.Text = sfd.FileName;
                    if (MessageBox.Show("CSVファイル名を自動設定しますか？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        txtOutputCsv.Text = Path.ChangeExtension(sfd.FileName, ".csv");
                    }
                }
            }
        }

        private void BtnSelectOutputCsv_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog { Filter = "CSVファイル|*.csv" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtOutputCsv.Text = sfd.FileName;
                }
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(txtInputFolder.Text))
                {
                    MessageBox.Show("入力フォルダが存在しません。", "エラー");
                    return;
                }

                var wavFiles = Directory.GetFiles(txtInputFolder.Text, "*.wav").OrderBy(f => f).ToList();

                // 先頭3文字のインデックス番号の重複チェック
                var indexGroups = wavFiles
                    .Select(f => Path.GetFileName(f).Substring(0, 3))
                    .GroupBy(index => index)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

                if (indexGroups.Any())
                {
                    string duplicateIndices = string.Join(", ", indexGroups);
                    MessageBox.Show($"インデックス番号の重複が検出されました: {duplicateIndices}。", "エラー");

                    // クリーンアップの確認
                    if (MessageBox.Show("入力フォルダをクリーンアップしますか？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        CleanupInputFolder();
                        MessageBox.Show("入力フォルダがクリーンアップされました。", "情報"); // 追加: メッセージ表示
                    }
                    return; // クリーンアップ後または「いいえ」の場合、処理を終了
                }

                if (!wavFiles.Any())
                {
                    MessageBox.Show("入力フォルダ内にWAVファイルが存在しません。", "エラー");
                    return;
                }

                if (string.IsNullOrEmpty(txtOutputWav.Text) || string.IsNullOrEmpty(txtOutputCsv.Text))
                {
                    MessageBox.Show("出力ファイル名またはCSVファイル名が指定されていません。", "エラー");
                    return;
                }

                ProcessWavFiles(wavFiles, txtOutputWav.Text, txtOutputCsv.Text, txtVoiceActor.Text);

                // ====================== /auto モード時は1秒後に自動終了 ======================
                var args = Environment.GetCommandLineArgs();
                if (args.Any(arg => arg.Equals("/auto", StringComparison.OrdinalIgnoreCase)))
                {
                    System.Threading.Tasks.Task.Delay(1000).ContinueWith(t =>
                    {
                        this.Invoke(new Action(() =>
                        {
                            autoSuccess = true;
                            this.Close();   // フォームを閉じる
                        }));
                    });
                }
                // ============================================================================
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー");
            }
        }

        private void BtnCleanup_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(txtInputFolder.Text))
                {
                    MessageBox.Show("入力フォルダが存在しません。", "エラー");
                    return;
                }

                if (MessageBox.Show("入力フォルダをクリーンアップしますか？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    CleanupInputFolder();
                    MessageBox.Show("入力フォルダがクリーンアップされました。", "情報");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリーンアップ中にエラーが発生しました: {ex.Message}", "エラー");
            }
        }

        private void CleanupInputFolder()
        {
            var existingWavFiles = Directory.GetFiles(txtInputFolder.Text, "*.wav");
            foreach (var file in existingWavFiles)
            {
                File.Delete(file);
            }
        }

        private void ProcessWavFiles(List<string> wavFiles, string outputWav, string outputCsv, string voiceActor)
        {
            var result = new StringBuilder();
            var csvLines = new List<string> { "from,to" };
            long totalSamples = 0;
            WaveFormat waveFormat = null;

            txtResult.Text = "";

            // ====================== ASS自動生成準備 ======================
            string assPath = Path.ChangeExtension(outputCsv, ".ass");
            string subtitleTxtPath = Path.Combine(txtInputFolder.Text, Path.GetFileName(txtInputFolder.Text) + ".txt");
            bool generateAss = chkAutoAss.Checked;
            List<string> subtitleLines = null;
            Parser budouxParser = null;
            int maxLineWidth = 24;        // デフォルト値
            string assHeader = "";
            int multiLineCount = 0;

            if (generateAss)
            {
                bool file_exists = File.Exists(subtitleTxtPath);
                if (!file_exists)
                {
                    string[] txtFiles = Directory.GetFiles(txtInputFolder.Text, "*.txt");

                    if (txtFiles.Length > 0)
                    {
                        file_exists = true;
                        subtitleTxtPath = txtFiles[0];  // 既定のファイルが見つからない時のリカバリー、テキストファイルの最初の1つを候補にする
                        Console.WriteLine("最初の.txtファイル: " + subtitleTxtPath);
                    }
                }

                if (!file_exists)
                {
                    if (MessageBox.Show($"{subtitleTxtPath}\nが存在しません。\nASSファイル自動生成をスキップしますか？",
                            "ASSファイル自動生成エラー", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == DialogResult.No)
                    {
                        return;
                    }
                    generateAss = false;
                }
                else
                {
                    subtitleLines = File.ReadAllLines(subtitleTxtPath, Encoding.UTF8).ToList();

                    // WAVファイル名とセリフテキストファイルを簡易照合
                    for (int i = 0; i < wavFiles.Count ; i++)
                    {
                        string wavFileName = Path.GetFileName(wavFiles[i]);

                        // WAVファイル名から [xxx] 部分を除去して比較用文字列を作成
                        string wavClean = System.Text.RegularExpressions.Regex.Replace(wavFileName, @"\[.*?\]", "");
                        wavClean = wavClean.Replace("...", "").Trim();

                        string txtLine = (i < subtitleLines.Count) ? subtitleLines[i] : "[no more subtitle lines...]";

                        // セリフファイルは「声優名,セリフ」なので声優名部分だけ抽出
                        int commaPos = txtLine.IndexOf(',');
                        string txtVoicePart = (commaPos >= 0 ? txtLine.Substring(0, commaPos) : txtLine).Trim();

                        if (wavClean.IndexOf(txtVoicePart, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            DialogResult res = MessageBox.Show(
                                $"ベリファイ失敗: 行 {i + 1}\n\n" +
                                $"WAV: {wavFileName}\n" +
                                $"TXT: {txtLine}\n\n" +
                                $"WAVファイルとセリフファイルが一致しません。\n" +
                                $"ASSファイル自動生成をスキップしますか？",
                                "ベリファイエラー",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button2);

                            if (res == DialogResult.No)
                            {
                                return; // 処理中断
                            }
                            else
                            {
                                generateAss = false;
                                break;
                            }
                        }
                    }

                    // BudouX初期化
                    string modelPath = Path.Combine(Application.StartupPath, "ja.json");
                    if (File.Exists(modelPath))
                    {
                        budouxParser = Parser.LoadDefaultJapaneseParser(modelPath);
                    }
                    else
                    {
                        MessageBox.Show("ja.jsonが見つかりません。ASS生成をスキップします。", "警告");
                        generateAss = false;
                    }

                    // ==================== ASS_header.txt を1回だけ読み込み ====================
                    string headerPath = Path.Combine(Application.StartupPath, "ASS_header.txt");
                    if (File.Exists(headerPath))
                    {
                        assHeader = File.ReadAllText(headerPath, Encoding.UTF8);

                        // MaxLineWidth を抽出
                        var match = System.Text.RegularExpressions.Regex.Match(assHeader,
                            @"MaxLineWidth\s*:\s*(\d+)",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        if (match.Success && int.TryParse(match.Groups[1].Value, out int w))
                        {
                            maxLineWidth = w;
                            // デバッグ用（必要なければ削除可）
                            // print_textbox($"MaxLineWidth: {maxLineWidth}文字に設定");
                        }
                    }
                    else
                    {
                        MessageBox.Show("ASS_header.txtが見つかりません。", "エラー");
                        generateAss = false;
                    }
                    // ==========================================================================
                }
            }
            // ============================================================

            using (var firstReader = new WaveFileReader(wavFiles[0]))
            {
                waveFormat = firstReader.WaveFormat;
            }

            int maxFileNameLength = wavFiles
                .Select(f => Path.GetFileName(f))
                .Max(f => f.Sum(c => IsFullWidth(c) ? 2 : 1));

            using (var output = new WaveFileWriter(outputWav, waveFormat))
            {
                for (int i = 0; i < wavFiles.Count; i++)
                {
                    var wavFile = wavFiles[i];

                    using (var reader = new WaveFileReader(wavFile))
                    {
                        if (reader.WaveFormat.SampleRate != waveFormat.SampleRate)
                        {
                            throw new InvalidOperationException("すべてのWAVファイルのサンプルレートが一致する必要があります。");
                        }

                        long sampleCount = reader.Length / reader.WaveFormat.BlockAlign;
                        double duration = (double)sampleCount / waveFormat.SampleRate;
                        double from = (double)totalSamples / waveFormat.SampleRate;
                        double to = from + duration;

                        string fileName = Path.GetFileName(wavFile);
                        int displayLength = fileName.Sum(c => IsFullWidth(c) ? 2 : 1);
                        int paddingLength = maxFileNameLength - displayLength;
                        string padding = new string(' ', paddingLength);

                        string marker = fileName.Contains(voiceActor) ? "(★)" : "";
                        string fromPadded = $"{from:F3}".PadLeft(8);
                        string toPadded = $"{to:F3}".PadLeft(8);
                        result.AppendLine($"{fileName}{padding} {sampleCount,10} samples {fromPadded}-{toPadded} sec {marker}");

                        if (marker != "")
                        {
                            csvLines.Add($"{from:F3},{to:F3}");
                        }

                        // ====================== ASS Dialogue追加 ======================
                        if (generateAss && subtitleLines != null && i < subtitleLines.Count)
                        {
                            string rawLine = subtitleLines[i];
                            int commaIndex = rawLine.IndexOf(',');
                            string subtitleText = (commaIndex >= 0 ? rawLine.Substring(commaIndex + 1) : rawLine).Trim();

                            if (!string.IsNullOrWhiteSpace(subtitleText))
                            {
                                // 「、」が1つだけあり、かつ前半・後半が24文字以内に収まる場合は優先改行
                                subtitleText = ApplyCommaPriorityWrap(subtitleText, maxLineWidth);

                                // 通常のBudouX処理（上記の優先改行が適用されなかった場合）
                                if (budouxParser != null && !subtitleText.Contains("\\N"))
                                {
                                    var chunks = budouxParser.Parse(subtitleText);
                                    subtitleText = WrapForAss(chunks, maxLineWidth);
                                }

                                // 3行超のチェック（\Nの個数で判定）
                                int lineCount = subtitleText.Split(new[] { "\\N" }, StringSplitOptions.None).Length;
                                if (lineCount > 3)
                                {
                                    multiLineCount++;
                                }

                                // ====================== キャラ別スタイル適用 ======================
                                string styleName = (marker == "") ? "Style1" : "Style2";
                                string dialogue = $"Dialogue: 0,{TimeSpanToAss(from)},{TimeSpanToAss(to)},{styleName},,0,0,0,,{subtitleText}";
                                assHeader += dialogue + "\r\n";
                            }
                        }
                        // ============================================================

                        byte[] buffer = new byte[1024];
                        int bytesRead;
                        while ((bytesRead = reader.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            output.Write(buffer, 0, bytesRead);
                        }

                        totalSamples += sampleCount;
                    }
                }
            }

            File.WriteAllLines(outputCsv, csvLines, Encoding.UTF8);

            // ASSファイル保存
            if (generateAss)
            {
                File.WriteAllText(assPath, assHeader, Encoding.UTF8);
            }

            double totalDuration = (double)totalSamples / waveFormat.SampleRate;
            int minutes = (int)(totalDuration / 60);
            double seconds = totalDuration % 60;
            result.AppendLine();
            result.AppendLine($"出力ファイル: {outputWav}");
            result.AppendLine($"CSVファイル: {outputCsv}");
            if (generateAss)
            {
                result.AppendLine($"ASSファイル: {assPath}");
                // 3行超警告
                if (multiLineCount > 0)
                {
                    MessageBox.Show(
                        $"ASSファイル自動生成\n\n" +
                        $"三行を超えるセリフが {multiLineCount} 件ありました。\n" +
                        $"ASSファイルを確認してください。\n\n" +
                        $"({assPath})",
                        "ASSファイル自動生成 - 警告",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }

            result.AppendLine($"出力ファイルサンプル数: {totalSamples}");
            result.AppendLine($"出力ファイル再生時間: {minutes:D3}:{seconds:00.000}");

            if (chkAutoCleanup.Checked)
            {
                CleanupInputFolder();
                result.AppendLine("(自動クリーンアップしました)");
            }

            txtResult.Text = result.ToString();
            txtResult.SelectionStart = txtResult.Text.Length;
            txtResult.ScrollToCaret();
        }

        // 「、」が1つだけあり、前半・後半がmaxWidth以内に収まる場合に改行
        private string ApplyCommaPriorityWrap(string text, int maxFullWidth = 24)
        {
            int maxWidth = maxFullWidth * 2;
            // 「、」の個数をカウント
            int commaCount = text.Count(c => c == '、');
            if (commaCount != 1)
                return text;

            if (GetFullWidthLength(text) < maxWidth)
                return text;

            int commaPos = text.IndexOf('、');
            string part1 = text.Substring(0, commaPos + 1);  // 「、」を含める
            string part2 = text.Substring(commaPos + 1);

            int len1 = GetFullWidthLength(part1);
            int len2 = GetFullWidthLength(part2);

            // 両方が24文字以内に収まる場合のみ優先改行
            if (len1 <= maxWidth && len2 <= maxWidth)
            {
                return part1 + "\\N" + part2;
            }

            return text;
        }

        // ====================== 自動改行関数（全角基準） ======================
        private string WrapForAss(List<string> chunks, int maxFullWidthChars = 24)
        {
            var sb = new StringBuilder();
            var line = new StringBuilder();
            int currentWidth = 0;                    // 半角換算の現在幅
            int maxWidth = maxFullWidthChars * 2;    // 全角24文字 = 半角48

            foreach (var chunk in chunks)
            {
                int chunkWidth = GetFullWidthLength(chunk);

                // 超過しそうなら改行
                if (currentWidth + chunkWidth > maxWidth && currentWidth > 0)
                {
                    if (sb.Length > 0) sb.Append("\\N");
                    sb.Append(line);
                    line.Clear();
                    currentWidth = 0;
                }

                line.Append(chunk);
                currentWidth += chunkWidth;
            }

            // 最後の行
            if (line.Length > 0)
            {
                if (sb.Length > 0) sb.Append("\\N");
                sb.Append(line);
            }

            return sb.ToString();
        }

        // 全角・半角を正確にカウント
        private int GetFullWidthLength(string text)
        {
            int length = 0;
            foreach (char c in text)
            {
                length += (c > '\u007F' || c == '　') ? 2 : 1;
            }
            return length;
        }

        private string TimeSpanToAss(double seconds)
        {
            var ts = TimeSpan.FromSeconds(seconds);
            return $"{(int)ts.TotalHours:D1}:{ts.Minutes:D2}:{ts.Seconds:D2}.{ts.Milliseconds / 10:D2}";
        }

        private bool IsFullWidth(char c)
        {
            return c > '\u007F';
        }
    }

    [DataContract]
    public class Settings
    {
        [DataMember]
        public string InputFolder { get; set; }
        [DataMember]
        public string OutputWav { get; set; }
        [DataMember]
        public string OutputCsv { get; set; }
        [DataMember]
        public string VoiceActor { get; set; }
    }
}