using System;
using System.Drawing;
using System.Windows.Forms;

namespace DeskChanger
{
    public partial class Dialog : Form
    {
        public Form1 mainForm;
        private int swapNum = -1;

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
            if (Control.ModifierKeys == Keys.Control && swapNum == -1)
            {

                var lamp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
                var g = Graphics.FromImage(lamp);
                g.FillEllipse(Brushes.Green, 1, 1, pictureBox1.Width - 2, pictureBox1.Height - 2);
                g.Dispose();
                pictureBox1.Image = lamp;

                swapNum = (int)numericUpDown1.Value;
                return;
            }

            numericUpDown1.Enabled = false;

            try
            {
                mainForm.button1_Click(swapNum);
            }
            catch (Exception)
            {
                numericUpDown1.Enabled = true;
            }

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            mainForm.classNum = (int)numericUpDown1.Value;
        }
    }
}
