
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;


namespace OutlookAddIn1
{


    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {

        private Office.IRibbonUI ribbon;
        public void OnTextButton(Office.IRibbonControl control)
        {
            MessageBox.Show("You clicked a different control." + control.Id);
        }
        public void OnTableButton(Office.IRibbonControl control)
        {
            

            try
            {
                string value = "Document 1";
                if (InputBox("New document", "New document name:", ref value) == DialogResult.OK)
                {
                    Outlook.Application oApp = new Outlook.Application();
                    Outlook.NameSpace oNS = oApp.GetNamespace("mapi");
                    Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Outlook.Items oItems = oInbox.Items;

                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelApp.Workbooks.Add();
                    Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;


                    workSheet.Cells[1, 1] = "Category";
                    workSheet.Cells[1, 2] = "TIME: " + DateTime.Now.ToLongTimeString();
                    workSheet.Cells[1, 3] = "Subject";
                    workSheet.Cells[1, 4] = "Date-Time";
                    workSheet.Cells[1, 5] = "SenderName";

                    var row = 1;
                    Outlook.MailItem newEmail = null;
                    foreach (object collectionItem in oItems)
                    {

                        newEmail = collectionItem as Outlook.MailItem;
                        if (newEmail != null)
                        {
                            //row++;
                            if (newEmail.Categories == "Orange Category")
                            {
                                row++;
                                workSheet.Cells[row, 1] = newEmail.Categories;
                                workSheet.Cells[row, 3] = newEmail.Subject;
                                workSheet.Cells[row, 4] = newEmail.ReceivedTime;
                                workSheet.Cells[row, 5] = newEmail.SenderName;
                            }
                            if (newEmail.Categories == "Green Category")
                            {
                                row++;
                                workSheet.Cells[row, 1] = newEmail.Categories;
                                workSheet.Cells[row, 3] = newEmail.Subject;
                                workSheet.Cells[row, 4] = newEmail.ReceivedTime;
                                workSheet.Cells[row, 5] = newEmail.SenderName;
                            }
                            else { }
                        }

                    }
                    workSheet.SaveAs(value);
                    MessageBox.Show("Your raport is saved in: " + value);
                }
                else
                {
                    MessageBox.Show("Operation cannceled");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public Ribbon1()
        {


        }
        public static DialogResult InputBox(string title, string promptText, ref string value)
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
            buttonOk.SetBounds(300, 100, 75, 23);
            buttonCancel.SetBounds(150, 100, 75, 23);

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
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {

                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private static void getAllEmails(Outlook.Items oItems)
        {
            String c = "";
            Outlook.MailItem newEmail = null;
            foreach (object collectionItem in oItems)
            {
                newEmail = collectionItem as Outlook.MailItem;
                if (newEmail != null)
                {
                    c += "\n" + newEmail.ReceivedTime + "  " + newEmail.SenderName;
                }
            }
            MessageBox.Show(c);
        }

        #endregion
    }
}
