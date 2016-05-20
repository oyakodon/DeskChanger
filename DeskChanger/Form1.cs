using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace DeskChanger
{
    public partial class Form1 : Form
    {
        public int classNum = 1;

        private List<int> rndList;
        private List<bool> checkedList = new List<bool>();
        private bool isLabelInited = false;
        
        public Form1()
        {
            InitializeComponent();
            var diag = new Dialog();
            diag.mainForm = this;
            diag.Show();

            // checkedListの初期化
            for (var i = 0; i < 48; i++)
            {
                checkedList.Add(false);
            }

            createCheckBox();
        }

        private void generateRndList()
        {
            rndList = new List<int>();

            for (var i = 0; i < classNum; ++i)
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

        private void swap(int k, int n)
        {
            var tmp = rndList[k];
            rndList[k] = rndList[n];
            rndList[n] = tmp;
        }

        private void createCheckBox()
        {
            var count = 0;
            for (var i = 0; i < 8; i++) // 横
            {
                for (var j = 0; j < 6; j++) // 縦
                {
                    var cb = new CheckBox();
                    cb.Name = "Check" + count.ToString();
                    cb.Height = 20;
                    cb.Width = 50;
                    cb.Top = 85 + i * 50;
                    cb.Left = 50 + j * 75;

                    Controls.Add(cb);
                    count++;
                }
            }
        }

        private void disposeCheckBox()
        {
            for (var i = 0; i < 48; i++)
            {
                var ctrlLst = this.Controls.Find("Check" + i.ToString(), true);
                checkedList[i] = ((CheckBox)ctrlLst[0]).Checked;
                this.Controls.Remove(ctrlLst[0]);
            }


            if (checkedList.Count(x => x) != 48 - classNum)
            {
                createCheckBox();
                throw new FormatException("空席の指定に誤りがあります。");
            }
            
        }

        public void button1_Click(int swapNum)
        {
            // 乱数配列生成
            generateRndList();

            // 不正はなかった
            if (swapNum != -1)
            {
                var swapNumIndex = rndList.IndexOf(swapNum);
                swap(swapNumIndex, classNum - 1);
            }

            // ラベルが生成されているか？
            if (isLabelInited)
            {
                // T : ラベルの値の変更
                rewriteLabels();

            } else
            {
                // F : ラベルの生成 -> ラベルの値の変更
                try
                {
                    disposeCheckBox();
                }

                catch (FormatException ex)
                {
                    MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    throw ex;
                }

                generateLabels();
            }

        }

        private void rewriteLabels()
        {
            var count = 0;

            for (var i = 0; i < 8; i++) // 横
            {
                for (var j = 0; j < 6; j++) // 縦
                {
                    var label = (Label)this.Controls.Find("deskLabel" + (6 * i + j).ToString(), true)[0];

                    if (checkedList[6 * i + j])
                    {
                        label.Text = "空";
                    } else
                    {
                        label.Text = rndList[count].ToString();
                        count++;
                    }
                }
            }
            
        }

        private void generateLabels()
        {
            var count = 0;

            for(var i = 0; i < 8; i++) // 横
            {
                for (var j = 0; j < 6; j++) // 縦
                {
                    var label = new Label();
                    if (checkedList[6 * i + j])
                    {
                        label.Text = "空";
                    } else
                    {
                        label.Text = rndList[count].ToString();
                        count++;
                    }
                    
                    label.Name = "deskLabel" + (6 * i + j).ToString();
                    label.Height = 20;
                    label.Width = 50;
                    label.Top = 80 + i * 50;
                    label.Left = 30 + j * 75;
                    label.TextAlign = ContentAlignment.BottomCenter;

                    Controls.Add(label);
                }
            }
            isLabelInited = true;

        }
    }
}
