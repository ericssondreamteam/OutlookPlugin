using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Debuger OurDebug = new Debuger();
        private bool DebugerOptymisation;
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
            oSheet.Cells[7, 13] = "In hands = ";

            if (row1 == 4) /* Gdy nie znajdzie zadnych maili w IN-HANDS */
                oSheet.Cells[7, 14].Value = 0;
            else
                oSheet.Cells[7, 14].Formula = "=ROWS(I5:F" + row1 + ")";
            if (row2 == 4) /* Gdy nie znajdzie zadnych maili w INFLOW */
                oSheet.Cells[5, 14].Value = 0;
            else
                oSheet.Cells[5, 14].Formula = "=ROWS(A5:A" + row2 + ")";
            if (row3 == 4) /* Gdy nie znajdzie zadnych maili w OUTFLOW */
                oSheet.Cells[6, 14].Value = 0;
            else
                oSheet.Cells[6, 14].Formula = "=ROWS(E5:E" + row3 + ")";
            oSheet.get_Range("N5", "N7").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        }

        public void insertDataExcel(Excel._Worksheet oSheet, int row, Outlook.MailItem newEmail, int amount, int whichCategory)
        {
            if (whichCategory == 1) //IN-HANDS
            {
                oSheet.Cells[row, 9] = newEmail.Subject;
                oSheet.Cells[row, 10] = amount;
                oSheet.Cells[row, 11] = newEmail.Categories;
            }
            if (whichCategory == 2) //INFLOW
            {
                oSheet.Cells[row, 1] = newEmail.Subject;
                oSheet.Cells[row, 2] = amount;
                oSheet.Cells[row, 3] = newEmail.Categories;
            }
            if (whichCategory == 3) //OUTFLOW
            {
                oSheet.Cells[row, 5] = newEmail.Subject;
                oSheet.Cells[row, 6] = amount;
                oSheet.Cells[row, 7] = newEmail.Categories;
            }

        }

        public DateTime getInflowDate()
        {
            DateTime today = GetFirstDayOfWeek(DateTime.Today);
            today = today.AddDays(-2).AddHours(5);
            return today;
        }
        public DateTime getTwoWeeksDate()
        {
            DateTime today = GetFirstDayOfWeek(DateTime.Today);
            today = today.AddHours(2);
            return today;
        }
        public int getConversationAmount(Outlook.MailItem newEmail)
        {
            Outlook.Conversation conv = newEmail.GetConversation();
            Outlook.Table table = conv.GetTable();
            return table.GetRowCount();
        }
        public int selectCorrectEmailType(Outlook.MailItem newEmail)
        {
            int typ = 0;
            if (newEmail.Categories == null) //inflow
            {
                //inflow
                if (getConversationAmount(newEmail) > 1) typ = 1;
                else typ = 2;
            }
            if (getConversationAmount(newEmail) > 1 && newEmail.ReceivedTime > getInflowDate()) //in hands
            {
                //in hands
                typ = 1;
            }
            else if (getConversationAmount(newEmail) == 1 && newEmail.ReceivedTime > getInflowDate()) //inflow
            {
                //inflow
                typ = 2;
            }
            else if ((newEmail.ReceivedTime > getInflowDate().AddDays(-7)) && (newEmail.ReceivedTime < getInflowDate())) //outflow
            {
                //outflow
                typ = 3;
            }
            return typ;
        }
        public void createCenterTables(Excel._Worksheet oSheet, int row1, int row2, int row3)
        {
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
        }


        public void OnTableButton(Office.IRibbonControl control)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            try
            {
                //Fajniejsza nazwa dla pliku raportu
                string OutputRaportFileName = "Raport " + DateTime.Now.ToString("dd/MM/yyyy");
                //Czy debugujemy
                if (Interaction.ShowDebugDialog("Debuger", "Turn on debuger?"))
                {
                    OurDebug.Enable();
                    DebugerOptymisation = true;
                }
                else
                    OurDebug.Disable();

                if (Interaction.SaveRaportDialog("New document", "New document name:", ref OutputRaportFileName) == DialogResult.OK)
                {
                    Outlook.Application oApp = new Outlook.Application();
                    NameSpace oNS = oApp.GetNamespace("mapi");
                    MAPIFolder oInbox2 = oApp.ActiveExplorer().CurrentFolder as MAPIFolder;
                    OurDebug.AppendInfo("Wybrany folder ", oInbox2.Name);
                    MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Items oItems = oInbox2.Items;
                    List<MailItem> emails = new List<MailItem>();

                    OurDebug.AppendInfo("Email's amount", oItems.Count.ToString());
                    oItems.Sort("[ReceivedTime]", true);//sortowanie od najnowszych

                    MailItem email1 = null;
                    int DebugForEachCounter = 0;
                    int DebugCorrectEmialsCounter = 0;
                    OurDebug.AppendInfo("\n\n ************************MAILS*******************\n\n");
                    foreach (object collectionItem in oItems)
                    {
                        try
                        {
                            DebugForEachCounter++;
                            email1 = collectionItem as MailItem;
                            if (email1 != null)
                            {
#if DebugerOptymisation
                                OurDebug.AppendInfo("Email  ",DebugCorrectEmialsCounter.ToString(), ": ", email1.Subject, email1.ReceivedTime.ToString());                                
#endif            
                                if (email1.ReceivedTime > getInflowDate().AddDays(-14))
                                {
                                    DebugCorrectEmialsCounter++;
                                    emails.Add(email1);
                                }
                                else
                                    break;
                            }
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Some error occured during first analysis\nIf You turn on debugger please go there");
                            OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "FIRST TRY CATCH\n", "Emial number:", DebugCorrectEmialsCounter.ToString(), "\n", e.Message, "\n", e.StackTrace);
                        }
                    }

                    OurDebug.AppendInfo("\n\n", "Ile razy foreach: ", DebugForEachCounter.ToString(), "Maile brane pod uwage po wstepnej selekcji: ", DebugCorrectEmialsCounter.ToString(), "\n\n");
                    oXL = new Excel.Application();
                    oXL.Visible = false;
                    oWB = (oXL.Workbooks.Add(Missing.Value));
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    createExcelColumn(oSheet);

                    var row1 = 4;
                    var row2 = 4;
                    var row3 = 4;
                    emails = emails.Distinct().ToList();//czy to potrzbne? 

                    foreach (MailItem newEmail in emails)
                    {
#if DebugerOptymisation
                        OurDebug.AppendInfo("Przed odczytem kategorii:",newEmail.Subject,newEmail.Categories, newEmail.ReceivedTime.ToString());
#endif
                        var typ = 0;
                        if (isMultipleCategoriesAndAnyOfTheireInterestedUs(newEmail.Categories))
                        {
#if DebugerOptymisation
                            OurDebug.AppendInfo("Po odczycie kategorii:",newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());
#endif
                            int emailConversationAmount = getConversationAmount(newEmail);
                            DateTime friday = getInflowDate();
                            typ = selectCorrectEmailType(newEmail);
#if DebugerOptymisation
                            OurDebug.AppendInfo("Nadany typ:",typ.ToString());
#endif
                            switch (typ)
                            {
                                case 1:
                                    row1++;
                                    insertDataExcel(oSheet, row1, newEmail, emailConversationAmount, 1);
                                    break;
                                case 2:
                                    row2++;
                                    insertDataExcel(oSheet, row2, newEmail, emailConversationAmount, 2);
                                    break;
                                case 3:
                                    row3++;
                                    insertDataExcel(oSheet, row3, newEmail, emailConversationAmount, 3);
                                    break;
                            }
                            oSheet.Columns.AutoFit();
                            oSheet.Cells[4, 1].EntireRow.Font.Bold = true;
                        }
                    }

                    createCenterTables(oSheet, row1, row2, row3);
                    createExcelSumCategories(oSheet, row1, row2, row3);
                    oWB.SaveAs(OutputRaportFileName, Excel.XlFileFormat.xlOpenXMLStrictWorkbook);
                    oWB.Close(true);
                    oXL.Quit();
                    Marshal.ReleaseComObject(oXL);
                    MessageBox.Show("Your raport is saved in: " + OutputRaportFileName);
                    OurDebug.AppendInfo("Your raport is SAVED :D");
                }
                else
                {
                    MessageBox.Show("Operation cannceled");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Some error occured during second analysis\nIf You turn on debugger please go there");
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "SECOND TRY CATCH\n", e.Message, "\n", e.StackTrace);
            }
            finally
            {
                if (OurDebug.IsEnable())
                {
                    OurDebug.SaveDebugInfoToFile(@"C:\Users\Public\DebugInfoRaportPlugin.txt");
                    MessageBox.Show("Plik debugowania zapisany w C:\\Users\\Public\nPlik: DebugInfoRaportPlugin.txt");
                }
            }
        }

        bool isMultipleCategoriesAndAnyOfTheireInterestedUs(string categories)
        {
            if (categories is null)
            {
                return true;
            }
            else
            {
                categories = categories.Trim();
                categories = categories.Replace(" ", "");
                categories.ToLower();
                string[] categoriesList = categories.Split(',');
                foreach (var cat in categoriesList)
                {   //No Response Necessary    or    Unknown     No Response Necessary, Unknown
                    if (!cat.Equals("noresponsenecessary") && !cat.Equals("unknown"))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public Ribbon1()
        {
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
