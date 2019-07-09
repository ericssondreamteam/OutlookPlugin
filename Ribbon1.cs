﻿
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

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
        public static DateTime GetFirstDayOfWeek(DateTime dayInWeek)
        {
            CultureInfo defaultCultureInfo = CultureInfo.CurrentCulture;
            return GetFirstDayOfWeek(dayInWeek, defaultCultureInfo);
        }
        public static DateTime GetFirstDayOfWeek(DateTime dayInWeek, CultureInfo cultureInfo)
        {
            DayOfWeek firstDay = cultureInfo.DateTimeFormat.FirstDayOfWeek;
            DateTime firstDayInWeek = dayInWeek.Date;
            while (firstDayInWeek.DayOfWeek != firstDay)
                firstDayInWeek = firstDayInWeek.AddDays(-1);

            return firstDayInWeek;
        }
        public void OnTableButton(Office.IRibbonControl control)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try
            {
                string value = "Document 1";
                if (InputBox("New document", "New document name:", ref value) == DialogResult.OK)
                {
                    Outlook.Application oApp = new Outlook.Application();
                    Outlook.NameSpace oNS = oApp.GetNamespace("mapi");
                    Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Outlook.Items oItems = oInbox.Items;


                    oXL = new Excel.Application();
                    oXL.Visible = false;
                    oWB = (oXL.Workbooks.Add(Missing.Value));
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    oSheet.Cells[1, 1] = "Raport Time: " + DateTime.Now.ToLongTimeString();
                    oSheet.Cells[1, 2] = "Raport Date: " + DateTime.Now.ToLongDateString();
                    oSheet.Cells[2, 2] = "Subject";
                    oSheet.Cells[2, 3] = "Count";
                    oSheet.Cells[2, 4] = "Inflow";
                    oSheet.Cells[2, 5] = "Outflow";
                    oSheet.Cells[2, 6] = "Inhance";
                    oSheet.Cells[2, 7] = "Category";
                    var row = 2;
                    Outlook.MailItem newEmail = null;
                    foreach (object collectionItem in oItems)
                    {

                        newEmail = collectionItem as Outlook.MailItem;
                        if (newEmail != null)
                        {
                            var typ = 0;
                            if (newEmail != null)
                            {
                                if (newEmail.Categories != null)
                                {
                                    //INFLOW
                                    DateTime today = GetFirstDayOfWeek(DateTime.Today);
                                    today = today.AddDays(-2).AddHours(5);
                                    //INHENCE
                                    Outlook.Conversation conv = newEmail.GetConversation();
                                    Outlook.SimpleItems items = conv.GetChildren(newEmail);
                                    Outlook.Table table = conv.GetTable();
                                    if (table.GetRowCount() > 1 && newEmail.ReceivedTime > today)
                                    {
                                        typ = 1;
                                    }
                                    else if (newEmail.ReceivedTime > today)
                                    {
                                        typ = 2;
                                    }
                                    else if ((newEmail.ReceivedTime > today.AddDays(-7)) && (newEmail.ReceivedTime < today))
                                    {
                                        typ = 3;
                                    }
                                }
                            }

                            Outlook.Conversation conv_ = newEmail.GetConversation();
                            Outlook.SimpleItems items_ = conv_.GetChildren(newEmail);
                            Outlook.Table table_ = conv_.GetTable();
                            switch(typ)
                            {
                                case 1:
                                    row++;
                                    oSheet.Cells[row, 2] = newEmail.Subject;
                                    oSheet.Cells[row, 3] = table_.GetRowCount();
                                    oSheet.Cells[row, 6].Value = 1;
                                    oSheet.Cells[row, 7] = newEmail.Categories;
                                    break;
                                case 2:
                                    row++;
                                    oSheet.Cells[row, 2] = newEmail.Subject;
                                    oSheet.Cells[row, 3] = table_.GetRowCount();
                                    oSheet.Cells[row, 4].Value = 1;
                                    oSheet.Cells[row, 7] = newEmail.Categories;
                                    break;
                                case 3:
                                    row++;
                                    oSheet.Cells[row, 2] = newEmail.Subject;
                                    oSheet.Cells[row, 3] = table_.GetRowCount();
                                    oSheet.Cells[row, 5].Value = 1;
                                    oSheet.Cells[row, 7] = newEmail.Categories;
                                    break;
                            }
                            oSheet.Columns.AutoFit();
                            oSheet.Cells[2, 2].EntireRow.Font.Bold = true;

                        }
                    }
                    oSheet.Cells[row + 3, 3] = "SUM";
                    oSheet.Cells[row + 4, 4] = "Inflow";
                    oSheet.Cells[row + 4, 5] = "Outflow";
                    oSheet.Cells[row + 4, 6] = "Inhence";
                    oSheet.Cells[row + 3, 4].Formula = "=SUM(D3:D" + row + ")";
                    oSheet.Cells[row + 3, 5].Formula = "=SUM(E3:E" + row + ")";
                    oSheet.Cells[row + 3, 6].Formula = "=SUM(F3:F" + row + ")";
                    oSheet.Cells[row + 3, 4].EntireRow.Font.Bold = true;

                    oWB.SaveAs(value, Excel.XlFileFormat.xlOpenXMLStrictWorkbook);
                    oWB.Close(true);
                    oXL.Quit();
                    Marshal.ReleaseComObject(oXL);
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
