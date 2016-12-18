using ClosedXML.Excel;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace DeskChanger
{
    public partial class Preference : Form
    {
        public Preference()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 名簿を使用するか
        /// </summary>
        public static bool useRecord { get; set; }

        /// <summary>
        /// 学科・クラス名
        /// </summary>
        public static string className { get; set; }

        /// <summary>
        /// クラスの人数
        /// </summary>
        public static int classNum { get; set; }

        /// <summary>
        /// 名簿
        /// </summary>
        public static Dictionary<int, string> record { get; set; }

        /// <summary>
        /// 前にしたい人
        /// </summary>
        public static List<int> forwards { get; set; }

        /// <summary>
        /// 空席リスト
        /// </summary>
        public static List<bool> emptySeats
        {
            get { return emptyChkBoxes.Select(x => x.Checked).ToList(); }
        }

        /// <summary>
        /// ウィザードの進捗状況
        /// </summary>
        private int progress { get; set; }

        /// <summary>
        /// 座席チェックボックス
        /// </summary>
        private static List<CheckBox> emptyChkBoxes { get; set; }

        /// <summary>
        /// フォーム読み込み時
        /// </summary>
        private void Preference_Load(object sender, EventArgs e)
        {
            // Startパネル
            // アイコンを埋め込みリソースから読み込んで表示
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            var bmp_appIco = new Bitmap(asm.GetManifestResourceStream("DeskChanger.deskChanger_icon.png"));
            var bmp_oykdnIco = new Bitmap(asm.GetManifestResourceStream("DeskChanger.oykdn_icon.png"));
            picbox_appicon.Image = bmp_appIco;
            picbox_oykdn.Image = bmp_oykdnIco;
            // バージョン情報の表示
            lblp1_1.Text = "DeskChanger " + asm.GetName().Version;

#if DEBUG
            numDebugProgress.Visible = true;
#endif

            // 変数
            useRecord = false;
            record = new Dictionary<int, string>();
            forwards = new List<int>();
            progress = 0;
            emptyChkBoxes = new List<CheckBox>();
        }

        /// <summary>
        /// Githubのリンククリック
        /// </summary>
        private void linklblp1_1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(linklblp1_1.Text);
        }

        /// <summary>
        /// 終了ボタン
        /// </summary>
        private void btnQuit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 開始 / 次へ押下時
        /// </summary>
        private void btnNext_Click(object sender, EventArgs e)
        {
            progress++;
            Debug.WriteLine("btnNext Clicked. Progress is " + progress);

            switch (progress)
            {
                case 1:
                    // Start -> PropTmpl
                    panelStart.Visible = false;
                    panelPropTmpl.Visible = true;
                    btnNext.Text = "次へ(&N)";
                    btnNext.Enabled = false;
                    lblStep1.Font = new Font(lblStep1.Font.FontFamily, 10.5f, FontStyle.Bold);

                    break;
                case 2:
                    className = tbClassName.Text;
                    classNum = (int)numClass.Value;

                    panelPropTmpl.Visible = false;
                    lblStep1.Font = new Font(lblStep1.Font.FontFamily, 9.0f, FontStyle.Regular);

                    if (useRecord)
                    {
                        // PropTmpl -> OpenXlsx
                        panelOpenXlsx.Visible = true;
                        btnNext.Enabled = false;
                        lblStep2.Font = new Font(lblStep2.Font.FontFamily, 10.5f, FontStyle.Bold);
                    } else
                    {
                        // PropTmpl -> ChkSeat
                        panelChkSeat.Visible = true;
                        lblStep3.Font = new Font(lblStep3.Font.FontFamily, 10.5f, FontStyle.Bold);
                        createCheckBox();
                        progress++;
                    }

                    break;
                case 3:
                    // OpenXlsx -> ChkSeat
                    panelOpenXlsx.Visible = false;
                    panelChkSeat.Visible = true;
                    lblStep2.Font = new Font(lblStep2.Font.FontFamily, 9.0f, FontStyle.Regular);
                    lblStep3.Font = new Font(lblStep3.Font.FontFamily, 10.5f, FontStyle.Bold);
                    createCheckBox();

                    break;
                case 4:
                    // ChkSeat -> Selection
                    if (48 - emptyChkBoxes.Count(x => x.Checked) != classNum)
                    {
                        MessageBox.Show("チェックされた数と人数との整合性がとれません。\nもう一度お確かめください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        progress--;
                        return;
                    }

                    panelChkSeat.Visible = false;
                    panelSelection.Visible = true;
                    lblStep3.Font = new Font(lblStep3.Font.FontFamily, 9.0f, FontStyle.Regular);
                    lblStep4.Font = new Font(lblStep4.Font.FontFamily, 10.5f, FontStyle.Bold);
                    btnNext.Text = "生成(&G)";

                    var dt = new DataTable();
                    dt.Columns.AddRange(new DataColumn[3] { new DataColumn("前", typeof(bool)), new DataColumn("出席番号", typeof(int)),
                    new DataColumn("氏名", typeof(string))});
                    if (useRecord)
                    {
                        foreach(var p in record)
                        {
                            dt.Rows.Add(false, p.Key, p.Value);
                        }
                    } else
                    {
                        for (var i = 0; i < classNum; i++)
                        {
                            dt.Rows.Add(false, i + 1, "");
                        }
                    }
                    dgvSelection.DataSource = dt;

                    foreach (DataGridViewColumn c in dgvSelection.Columns)
                    {
                        c.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }

                    break;
                case 5:
                    // Selection -> Generate
                    if (forwards.Count > 12)
                    {
                        MessageBox.Show("前にする人数が12名(6人×2列)よりも多く設定されています。\nもう一度お確かめください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        progress--;
                        return;
                    }

                    (new SeatVisualizer()).ShowDialog();
                    this.Close();

                    break;
                default: break;
            }

        }

        /// <summary>
        /// チェックボックスの生成
        /// </summary>
        private void createCheckBox()
        {
            const int x = 55;
            const int y = 65;

            for (var i = 0; i < 8; i++) // 横
            {
                for (var j = 0; j < 6; j++) // 縦
                {
                    var cb = new CheckBox();
                    cb.Name = "seatChk" + (i * 6 + j).ToString();
                    cb.Height = 20;
                    cb.Width = 20;
                    cb.Top = y + i * 35;
                    cb.Left = x + j * 45;

                    panelChkSeat.Controls.Add(cb);
                    emptyChkBoxes.Add(cb);
                }
            }
        }

        private void btnCpXlsx_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbClassName.Text))
            {
                MessageBox.Show("クラス名が入力されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var dlg = MessageBox.Show("クラス設定\n　クラス名："+ tbClassName.Text + "\n　人数：" + (int)numClass.Value + "\n\n入力に間違いはありませんか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlg != DialogResult.Yes) return;

            tbClassName.Enabled = false;
            numClass.Enabled = false;
            className = tbClassName.Text;
            classNum = (int)numClass.Value;

            // ファイルの選択
            var fileName = "";
            var sfd = new SaveFileDialog();
            sfd.FileName = "名簿_テンプレート.xlsx";
            sfd.InitialDirectory = Environment.CurrentDirectory;
            sfd.Filter = "Excelファイル(*.xlsx)|*.xlsx";
            sfd.Title = "テンプレートのコピー先のファイルを選択してください";
            sfd.RestoreDirectory = true;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                fileName = sfd.FileName;
            } else
            {
                MessageBox.Show("キャンセルされました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 名簿テンプレートの作成
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Sheet1");

                // セルの設定
                ws.Range("B2:C51").Style
                    .Border.SetTopBorder(XLBorderStyleValues.Thin)
                    .Border.SetBottomBorder(XLBorderStyleValues.Thin)
                    .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                    .Border.SetRightBorder(XLBorderStyleValues.Thin);

                ws.Column("C").Width = 14;

                var r = ws.Range("B2:C2");
                r.Merge();
                r.Value = className + "名簿";
                r.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Cell("B3").Value = "出席番号";
                ws.Cell("C3").Value = "氏名";

                for (var i = 1; i <= classNum; i++)
                {
                    ws.Cell(3 + i, 2).Value = i.ToString();
                    ws.Cell(3 + i, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                }

                wb.SaveAs(fileName);
                
            }
            catch (Exception ex)
            {
                var logPath = Environment.CurrentDirectory + "\\error.log";
                var log = "";
                if (System.IO.File.Exists(logPath)) log = System.IO.File.ReadAllText(logPath);
                log = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace + Environment.NewLine + log;
                System.IO.File.WriteAllText(logPath, log);
                MessageBox.Show("予期しないエラーが発生しました。\n詳しくは開発者にお問い合わせください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("テンプレートのコピーが完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            btnCpXlsx.Enabled = false;
            chkUseRec.Enabled = false;
        }

        private void btnOpenXlsx_Click(object sender, EventArgs e)
        {
            try
            {
                // ファイルの読み込み
                var ofd = new OpenFileDialog();
                ofd.InitialDirectory = Environment.CurrentDirectory;
                ofd.Filter = "Excelファイル(*.xlsx)|*.xlsx";
                ofd.Title = "名簿ファイルを選択してください";
                ofd.RestoreDirectory = true;

                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("キャンセルされました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 名簿の読み込み
                using (var wb = new XLWorkbook(ofd.FileName))
                {
                    var ws = wb.Worksheet("Sheet1");
                    var r = ws.Range("B4:B51").CellsUsed();
                    foreach (var cell in r)
                    {
                        record.Add(cell.Value.CastTo<int>(), cell.CellRight().Value.CastTo<string>());
                    }

                }

            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("ファイルが使用中です。\nファイルを閉じてから再試行してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                var logPath = Environment.CurrentDirectory + "\\error.log";
                var log = "";
                if (System.IO.File.Exists(logPath)) log = System.IO.File.ReadAllText(logPath);
                log = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace + Environment.NewLine + log;
                System.IO.File.WriteAllText(logPath, log);
                MessageBox.Show("予期しないエラーが発生しました。\n詳しくは開発者にお問い合わせください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[2] { new DataColumn("出席番号", typeof(int)),
            new DataColumn("氏名", typeof(string))});
            foreach(var p in record)
            {
                dt.Rows.Add(p.Key, p.Value);
            }
            dgvRecord.DataSource = dt;

            foreach (DataGridViewColumn c in dgvRecord.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            lblp3_2.Text = "人数 : " + record.Count;
            classNum = record.Count;
            btnNext.Enabled = true;
        }

        private void dgvSelection_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (0 <= e.RowIndex && e.RowIndex <= classNum)
            {
                if (dgvSelection.Rows[e.RowIndex].Cells[0].Value.CastTo<bool>())
                {
                    forwards.Remove(dgvSelection.Rows[e.RowIndex].Cells[1].Value.CastTo<int>());
                    dgvSelection.Rows[e.RowIndex].Cells[0].Value = false;
                } else
                {
                    var dlg = MessageBox.Show((e.RowIndex + 1) + "番を前にしますか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dlg != DialogResult.OK) return;
                    forwards.Add(dgvSelection.Rows[e.RowIndex].Cells[1].Value.CastTo<int>());
                    dgvSelection.Rows[e.RowIndex].Cells[0].Value = true;
                }

                lblp5_2.Text = "前にする人数：" + forwards.Count;
            }
        }

        private void tbClassName_TextChanged(object sender, EventArgs e)
        {
            btnNext.Enabled = true;
        }

        private void chkUseRec_CheckedChanged(object sender, EventArgs e)
        {
            useRecord = chkUseRec.Checked;
            btnCpXlsx.Enabled = useRecord;
        }

        private void Preference_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (progress != 5)
            {
                var dlg = MessageBox.Show("ウィザードが進行中です。終了しますか？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg != DialogResult.OK)
                {
                    e.Cancel = true;
                }
            }
        }

        /// <summary>
        /// !デバッグモード!
        /// </summary>
        private void numDebugProgress_ValueChanged(object sender, EventArgs e)
        {
            btnNext.Enabled = false;

            progress = (int)numDebugProgress.Value;
            panelStart.Visible = false;
            panelPropTmpl.Visible = false;
            panelOpenXlsx.Visible = false;
            panelChkSeat.Visible = false;
            panelSelection.Visible = false;
            switch (progress)
            {
                case 0: panelStart.Visible = true; break;
                case 1: panelPropTmpl.Visible = true; break;
                case 2: panelOpenXlsx.Visible = true; break;
                case 3: createCheckBox(); panelChkSeat.Visible = true; break;
                case 4: createCheckBox(); panelSelection.Visible = true; break;
                default: break;
            }
        }
    }
}
