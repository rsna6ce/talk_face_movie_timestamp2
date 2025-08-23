using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NAudio.Wave;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

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
        private Button btnCleanup; // 新しいボタン

        private readonly string settingsFilePath = Path.Combine(Application.StartupPath, "settings.json");

        public Form1()
        {
            InitializeComponent();
            InitializeUI();
        }

        private void InitializeUI()
        {
            // フォーム設定（デザイナーで設定済みの場合はコメントアウト）
            // this.Text = "Talk Face Timestamp2";
            // this.Width = 800;
            // this.Height = 600;

            // 入力フォルダ
            Label lblInputFolder = new Label { Text = "入力フォルダ:", Location = new System.Drawing.Point(10, 10), Width = 100 };
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

            // クリーンアップボタン（新しく追加）
            btnCleanup = new Button { Text = "入力フォルダクリーンアップ", Location = new System.Drawing.Point(120, 130), Width = 150 };
            btnCleanup.Click += BtnCleanup_Click;

            // 結果表示
            Label lblResult = new Label { Text = "結果:", Location = new System.Drawing.Point(10, 160), Width = 100 };
            txtResult = new TextBox { Location = new System.Drawing.Point(10, 190), Width = 780, Height = 350, Multiline = true, ReadOnly = true, Font = new System.Drawing.Font("ＭＳ ゴシック", 9), ScrollBars = ScrollBars.Vertical };

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] { lblInputFolder, txtInputFolder, btnSelectInputFolder,
                                                  lblOutputWav, txtOutputWav, btnSelectOutputWav,
                                                  lblOutputCsv, txtOutputCsv, btnSelectOutputCsv,
                                                  lblVoiceActor, txtVoiceActor, btnStart, btnCleanup, // btnCleanupを追加
                                                  lblResult, txtResult });

            // イベントハンドラの登録
            this.Load += Form1_Load;
            this.FormClosing += Form1_FormClosing;
        }

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
                // 読み込み失敗時はデフォルト値のまま（エラーは無視）
                System.Diagnostics.Debug.WriteLine($"設定の読み込みに失敗: {ex.Message}");
            }
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
                // 保存失敗時はエラー表示（任意）
                System.Diagnostics.Debug.WriteLine($"設定の保存に失敗: {ex.Message}");
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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー");
            }
        }

        // 新しいクリーンアップボタンのイベントハンドラ
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
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリーンアップ中にエラーが発生しました: {ex.Message}", "エラー");
            }
        }

        // クリーンアップ処理を独立したメソッドに抽出
        private void CleanupInputFolder()
        {
            var existingWavFiles = Directory.GetFiles(txtInputFolder.Text, "*.wav");
            foreach (var file in existingWavFiles)
            {
                File.Delete(file);
            }
            MessageBox.Show("入力フォルダがクリーンアップされました。", "情報");
        }

        private void ProcessWavFiles(List<string> wavFiles, string outputWav, string outputCsv, string voiceActor)
        {
            var result = new StringBuilder();
            var csvLines = new List<string> { "from,to" };
            long totalSamples = 0;
            WaveFormat waveFormat = null;

            txtResult.Text = "";

            using (var firstReader = new WaveFileReader(wavFiles[0]))
            {
                waveFormat = firstReader.WaveFormat;
            }

            int maxFileNameLength = wavFiles
                .Select(f => Path.GetFileName(f))
                .Max(f => f.Sum(c => IsFullWidth(c) ? 2 : 1));

            using (var output = new WaveFileWriter(outputWav, waveFormat))
            {
                foreach (var wavFile in wavFiles)
                {
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

            double totalDuration = (double)totalSamples / waveFormat.SampleRate;
            int minutes = (int)(totalDuration / 60);
            double seconds = totalDuration % 60;
            result.AppendLine();
            result.AppendLine($"出力ファイル: {outputWav}");
            result.AppendLine($"CSVファイル: {outputCsv}");
            result.AppendLine($"出力ファイルサンプル数: {totalSamples}");
            result.AppendLine($"出力ファイル再生時間: {minutes:D3}:{seconds:00.000}");

            txtResult.Text = result.ToString();
            txtResult.SelectionStart = txtResult.Text.Length;
            txtResult.ScrollToCaret();
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