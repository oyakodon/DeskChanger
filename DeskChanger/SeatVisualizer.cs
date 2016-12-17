using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using ClosedXML.Excel;

namespace DeskChanger
{
    public partial class SeatVisualizer : Form
    {
        /// <summary>
        /// ランダムな配列 (=席替え後の座席表)
        /// </summary>
        private List<int> rndList { get; set; }
        
        /// <summary>
        /// 座席表が保存されたかどうか
        /// </summary>
        private bool isSaved { get; set; }

        public SeatVisualizer()
        {
            InitializeComponent();

            generateRndList();
            MoveForwards();
            generateLabels();

            this.Text = Preference.className + " 座席表 - DeskChanger";
        }

        /// <summary>
        /// ランダムな配列を作成します
        /// </summary>
        private void generateRndList()
        {
            rndList = new List<int>();

            for (var i = 0; i < Preference.classNum; ++i)
                rndList.Add(i + 1);

            var rnd = new Random();
            var n = rndList.Count;

            while (n > 1)
            {
                --n;
                var k = rnd.Next(n + 1);
                swap(k, n);
            }

        }

        /// <summary>
        /// rndListのk番目とn番目を交換します
        /// </summary>
        private void swap(int k, int n)
        {
            var tmp = rndList[k];
            rndList[k] = rndList[n];
            rndList[n] = tmp;
        }

        /// <summary>
        /// 前にしたい人を(配列内で)前に移動させる
        /// </summary>
        public void MoveForwards()
        {
            if (Preference.classNum >= 12 && Preference.forwards.Count <= 12)
            {
                var rnd = new Random();
                foreach (var sNum in Preference.forwards)
                {
                    var swapNumIndex = rndList.IndexOf(sNum);
                    var forward = new List<int>(rndList.Take(12).ToList());
                    forward.RemoveAll(x => Preference.forwards.Contains(x));
                    var forwardIndex = rndList.IndexOf(forward[rnd.Next(forward.Count)]);
                    swap(swapNumIndex, forwardIndex);
                }
            }
        }

        /// <summary>
        /// 出席番号・氏名が書かれたラベルを作成します 
        /// </summary>
        private void generateLabels()
        {
            var count = 0;

            for (var i = 0; i < 8; i++) // 横
            {
                for (var j = 0; j < 6; j++) // 縦
                {
                    var label = new Label();
                    if (Preference.emptySeats[6 * i + j])
                    {
                        label.Text = "空";
                    } else
                    {
                        label.Text = rndList[count].ToString();
                        if (Preference.useRecord)
                        {
                            label.Text += Environment.NewLine + Preference.record[rndList[count] - 1];
                        }
                        count++;
                    }

                    label.Name = "deskLabel" + (6 * i + j).ToString();
                    label.Font = new Font("Meiryo UI", 13, FontStyle.Regular);
                    label.Height = 50;
                    label.Width = 135;
                    label.Top = 80 + i * 65;
                    label.Left = 30 + j * 120;
                    label.TextAlign = ContentAlignment.BottomCenter;

                    Controls.Add(label);
                }
            }
        }

        /// <summary>
        /// 座席表をExcelファイルに保存します
        /// </summary>
        private void SaveXlsx (object s, EventArgs e)
        {
            // ファイルの選択
            var fileName = "";
            var sfd = new SaveFileDialog();
            sfd.FileName = Preference.className + "座席表_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
            sfd.InitialDirectory = Environment.CurrentDirectory;
            sfd.Filter = "Excelファイル(*.xlsx)|*.xlsx";
            sfd.Title = "座席表の保存先のファイルを選択してください";
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
                //リソースを読み込む
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                var stream = asm.GetManifestResourceStream("DeskChanger.Template_SeatTable.xlsx");
                var buffer = new byte[stream.Length];
                stream.Read(buffer, 0, (int)buffer.Length);
                stream.Close();

                //ファイルに書き込む
                var fs = new System.IO.FileStream(fileName, System.IO.FileMode.Create);
                fs.Write(buffer, 0, (int)buffer.Length);
                fs.Close();
                
                // 座席情報の書き込み
                using (var wb = new XLWorkbook(fileName))
                {
                    var ws = wb.Worksheet("Sheet1");
                    ws.Cell("D2").Value = Preference.className;
                    ws.Cell("C15").Value = Preference.classNum;
                    ws.Cell("G15").Value = DateTime.Now.ToString("yyyy/MM/dd");

                    var count = 0;

                    for (var i = 0; i < 8; i++)
                    {
                        for (var j = 0; j < 6; j++)
                        {
                            if (Preference.emptySeats[6 * i + j])
                            {
                                // 空席
                                ws.Cell("C16").CopyTo(ws.Cell(6 + i, 2 + j));

                            } else
                            {
                                var str = rndList[count].ToString();
                                if (Preference.useRecord)
                                {
                                    str += Environment.NewLine + Preference.record[rndList[count] - 1];
                                }

                                ws.Cell(6 + i, 2 + j).Value = str;

                                count++;
                            }
                        }
                    }

                    wb.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("予期しないエラーが発生しました。\n詳しくは開発者にお問い合わせください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var appendText = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace;
                System.IO.File.AppendAllText(Environment.CurrentDirectory + "\\error.log", appendText);
                return;
            }

            MessageBox.Show("座席表の保存が完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            btnSave.Enabled = false;
            isSaved = true;
        }

        /// <summary>
        /// 終了時
        /// </summary>
        private void SeatVisualizer_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!isSaved)
            {
                var dlg = MessageBox.Show("座席表が保存されていません。終了しますか？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg != DialogResult.OK)
                {
                    e.Cancel = true;
                }
            }
        }
    }
}
