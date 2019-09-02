using System;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
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
            label1.Text = Loading.fullInfoBox;
        }

        private void Label1_Click(object sender, EventArgs e)
        {
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            
        }
    }
}
