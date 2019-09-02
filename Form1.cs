using System;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class Form1 : Form
    {
        public Form1(ref string OutputRaportFileName)
        {
            InitializeComponent();
            textBox2.Text = OutputRaportFileName;

            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 6000;
            toolTip1.InitialDelay = 300;
            toolTip1.ReshowDelay = 300;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip1.SetToolTip(textBox1, "e.g. NC Mailbox or Karol Lasek");
            toolTip1.SetToolTip(textBox2, "Please enter you report name.");
            toolTip1.SetToolTip(dateTimePicker1, "The report is created up to \nFriday after 5p.m. two weeks ago.");
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
            if(textBox2.Text.Length > 0)
            {
                Settings.OutputRaportFileName = textBox2.Text;
                if (DateTime.Parse(Settings.raportDate) > DateTime.Now)
                {
                    Settings.ifWeDoRaport = DialogResult.Cancel;
                    MessageBox.Show("No chyba nie...\nNie robimy raportów w przyszłości");
                }
                else
                {
                    Close();
                }
                
            }
            else
            {
                label5.Text = "You must fill this field.";
            }
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Settings.ifWeDoRaport = DialogResult.Cancel;
            Close();
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Settings.raportDate = dateTimePicker1.Value.ToString();
            Settings.OutputRaportFileName= "Raport_" + dateTimePicker1.Value.ToString("dd_MM_yyyy");
            textBox2.Text = Settings.OutputRaportFileName;
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            Settings.OutputRaportFileName = textBox2.Text;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void Label4_Click(object sender, EventArgs e)
        {

        }

        private void Label5_Click(object sender, EventArgs e)
        {

        }
    }
}
