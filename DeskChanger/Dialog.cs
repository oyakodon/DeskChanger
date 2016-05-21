using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace DeskChanger
{
    public partial class Dialog : Form
    {
        public Form1 mainForm;
        private List<int> swapNums = null;

        public Dialog()
        {
            InitializeComponent();
        }

        private void Dialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            mainForm.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (swapNums == null)
            {
                try
                {
                    swapNums = new List<int>();
                    addSwaps();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            numericUpDown1.Enabled = false;

            try
            {
                mainForm.button1_Click(swapNums);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                numericUpDown1.Enabled = true;
            }

        }

        private void addSwaps()
        {
            for(var i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                int swap;
                int.TryParse(dataGridView1[0, i].Value.ToString(), out swap);
                if(swap != 0 && swap <= 48)
                {
                    swapNums.Add(swap);

                } else
                {
                    swapNums = null;
                    throw new FormatException("表の数字に誤りがあります。");
                }
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            mainForm.classNum = (int)numericUpDown1.Value;
        }
    }
}
