
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
        public void createExcelColumn(Excel._Worksheet oSheet)
        {
            oSheet.Cells[1, 1] = "Raport Time: " + DateTime.Now.ToLongTimeString();
            oSheet.Cells[1, 2] = "Raport Date: " + DateTime.Now.ToLongDateString();

            oSheet.Cells[3, 1] = "INFLOW";
            oSheet.Cells[3, 5] = "OUTFLOW";
            oSheet.Cells[3, 9] = "IN-HANDS";


            oSheet.Cells[4, 1] = "Subject";
            oSheet.Cells[4, 2] = "Messages amount";
            oSheet.Cells[4, 3] = "Category";

            oSheet.Cells[4, 5] = "Subject";
            oSheet.Cells[4, 6] = "Messages amount";
            oSheet.Cells[4, 7] = "Category";

            oSheet.Cells[4, 9] = "Subject";
            oSheet.Cells[4, 10] = "Messages amount";
            oSheet.Cells[4, 11] = "Category";
        }

        public void createExcelSumCategories(Excel._Worksheet oSheet, int row1, int row2, int row3)
        {
            oSheet.Cells[4, 13] = "SUMMARY";
            oSheet.Cells[5, 13] = "Inflow  = ";
            oSheet.Cells[6, 13] = "Outflow = ";
            oSheet.Cells[7, 13] = "Inhence = ";
            oSheet.Cells[5, 14].Formula = "=ROWS(A5:A" + row2 + ")";
            oSheet.Cells[6, 14].Formula = "=ROWS(E5:E" + row3 + ")";
            oSheet.Cells[7, 14].Formula = "=ROWS(I5:F" + row1 + ")";
            oSheet.get_Range("N5", "N7").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        public void insertDataExcel(Excel._Worksheet oSheet, int row, Outlook.MailItem newEmail, Outlook.Table table_, int whichCategory)
        {
            if(whichCategory == 1) //IN-HANDS
            {
                oSheet.Cells[row, 9] = newEmail.Subject;
                oSheet.Cells[row, 10] = table_.GetRowCount();
                oSheet.Cells[row, 11] = newEmail.Categories;
            }
            if(whichCategory == 2) //INFLOW
            {
                oSheet.Cells[row, 1] = newEmail.Subject;
                oSheet.Cells[row, 2] = table_.GetRowCount();
                oSheet.Cells[row, 3] = newEmail.Categories;
            }
            if (whichCategory == 3) //OUTFLOW
            {
                oSheet.Cells[row, 5] = newEmail.Subject;
                oSheet.Cells[row, 6] = table_.GetRowCount();
                oSheet.Cells[row, 7] = newEmail.Categories;
            }

        }

        public DateTime getInflowDate()
        {
            DateTime today = GetFirstDayOfWeek(DateTime.Today);
            today = today.AddDays(-2).AddHours(5);
            return today;
        }

        public int getConversationAmount(Outlook.MailItem newEmail)
        {
            Outlook.Conversation conv = newEmail.GetConversation();
            Outlook.SimpleItems items = conv.GetChildren(newEmail);
            Outlook.Table table = conv.GetTable();
            return table.GetRowCount();
        }

        public int selectCorrectEmailType(Outlook.MailItem newEmail)
        {
            int typ = 0;
            if (getConversationAmount(newEmail) > 1 && newEmail.ReceivedTime > getInflowDate())
            {
                typ = 1;
            }
            else if (newEmail.ReceivedTime > getInflowDate())
            {
                typ = 2;
            }
            else if ((newEmail.ReceivedTime > getInflowDate().AddDays(-7)) && (newEmail.ReceivedTime < getInflowDate()))
            {
                typ = 3;
            }
            return typ;
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

                    createExcelColumn(oSheet);

                    var row1 = 4;
                    var row2 = 4;
                    var row3 = 4;
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
                                    DateTime friday = getInflowDate();
                                    int emailConversationAmount = getConversationAmount(newEmail);

                                    typ = selectCorrectEmailType(newEmail);

                                    Outlook.Conversation conv_ = newEmail.GetConversation();
                                    Outlook.SimpleItems items_ = conv_.GetChildren(newEmail);
                                    Outlook.Table table_ = conv_.GetTable();

                                    switch (typ)
                                    {
                                        case 1:
                                            row1++;
                                            insertDataExcel(oSheet, row1, newEmail, table_, 1);
                                            break;
                                        case 2:
                                            row2++;
                                            insertDataExcel(oSheet, row2, newEmail, table_, 2);
                                            break;
                                        case 3:
                                            row3++;
                                            insertDataExcel(oSheet, row3, newEmail, table_, 3);
                                            break;
                                    }
                                    oSheet.Columns.AutoFit();
                                    oSheet.Cells[4, 1].EntireRow.Font.Bold = true;
                                }
                            }
                        }
                    }
                    Excel.Range tRange1 = oSheet.get_Range("A4", "C" + row2);
                    oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange1,
                        Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "INFLOW";
                    oSheet.ListObjects["INFLOW"].TableStyle = "TableStyleMedium9";

                    Excel.Range tRange2 = oSheet.get_Range("E4", "G" + row3);
                    oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange2,
                        Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "OUTFLOW";
                    oSheet.ListObjects["OUTFLOW"].TableStyle = "TableStyleMedium12";

                    Excel.Range tRange3 = oSheet.get_Range("I4", "K" + row1);
                    oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange3,
                        Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "IN-HANDS";
                    oSheet.ListObjects["IN-HANDS"].TableStyle = "TableStyleMedium14";

                    oSheet.get_Range("B5", "B" + row2).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    oSheet.get_Range("F5", "F" + row3).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    oSheet.get_Range("J5", "J" + row1).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    createExcelSumCategories(oSheet, row1, row2, row3); //TRZEBA ZMIENIC
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
