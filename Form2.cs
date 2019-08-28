using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Process.Start(@path);
            Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.Text = Ribbon1.fullInfoBox;
        }

        private void Label1_Click(object sender, EventArgs e)
        {
            
        }


    }
}
