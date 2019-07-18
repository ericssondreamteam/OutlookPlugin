using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    class Interaction
    {
        public static DialogResult SaveRaportDialog(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;
            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 300, 13);
            textBox.SetBounds(12, 50, 400, 20);
            buttonOk.SetBounds(300, 100, 100, 30);
            buttonCancel.SetBounds(150, 100, 100, 30);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(424, 150);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        public static List<bool> ShowDebugDialog(string text, string textExcel, string textTxt, string caption)
        {
            List<bool> checkList = new List<bool>();
            checkList.Add(false);
            checkList.Add(false);
            checkList.Add(false);
            Form prompt = new Form();
            prompt.Width = 350;
            prompt.Height = 150;
            prompt.Text = caption;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            FlowLayoutPanel panel = new FlowLayoutPanel();

            CheckBox chk = new CheckBox();
            chk.Text = text;
            CheckBox chkExcel = new CheckBox();
            chkExcel.Text = textExcel;
            CheckBox chkTxt = new CheckBox();
            chkTxt.Text = textTxt;
            Button ok = new Button() { Text = "Confirm" };
            //ok.SetBounds(0, 200, 200, 30);
            ok.Click += (sender, e) => { prompt.Close(); };

            panel.Controls.Add(chk);
            panel.Controls.Add(chkExcel);
            panel.Controls.Add(chkTxt);
            panel.Controls.Add(ok);

            prompt.Controls.Add(panel);
            prompt.ShowDialog();

            if (chk.Checked)
                checkList[0] = true;
            if (chkExcel.Checked)
                checkList[1] = true;
            if (chkTxt.Checked)
                checkList[2] = true;
            return checkList;

        }
    }
}
