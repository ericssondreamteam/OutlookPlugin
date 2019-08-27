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
        public Form1(ref string OutputRaportFileName)
        {
            InitializeComponent();
            textBox2.Text = OutputRaportFileName;
        }



        private void Button1_Click(object sender, EventArgs e)
        {
            
            Settings.boxMailName = textBox1.Text;
            Settings.raportDate = dateTimePicker1.Value.ToString();
           
            if (checkBox1.Checked)
                Settings.checkList[1] = true;
            if (checkBox2.Checked)
                Settings.checkList[2] = true;
            if (checkBox3.Checked)
                Settings.checkList[0] = true;
            Settings.ifWeDoRaport = DialogResult.OK;
            Settings.OutputRaportFileName = textBox2.Text;
            Close();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Settings.ifWeDoRaport = DialogResult.Cancel;
        }

    }
}
