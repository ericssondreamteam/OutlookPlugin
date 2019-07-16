using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class Form1 : Form
    {
        public static int value;
        public Form1()
        {
            InitializeComponent();
            this.timer1.Start();
            this.progressBar1.Maximum = Ribbon1.counter;
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void ProgressBar1_Click(object sender, EventArgs e)
        {

        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Value = value;
            if(this.progressBar1.Value > Ribbon1.counter - 15)
            {
                //MessageBox.Show("ukryj sie");
                this.timer1.Enabled = false;
                this.Close();
            }
        }
        public static void incrementValue(int progress)
        {
            value = progress;
        }
    }
}
