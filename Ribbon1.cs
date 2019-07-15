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
        public DateTime getInflowDate()
        {
            DateTime today = GetFirstDayOfWeek(DateTime.Today);
            today = today.AddDays(-2).AddHours(5);
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

        public List<MailItem> emailsWithoutDuplicates(List<MailItem> emails)
        {
            

            
            for (int i = 0; i < emails.Count; i++)
            {
                for (int j = i + 1; j < emails.Count; j++)
                {
                    if (emails[i].ConversationID.Equals(emails[j].ConversationID))
                    {
                        OurDebug.AppendInfo("\n\nPorwannie: i:", emails[i].ConversationID, "j:", emails[j].ConversationID, "\n\n");
                        emails.RemoveAt(j);
                    }

                }
            }

            return emails;
        }

        public void OnTableButton(Office.IRibbonControl control)
        {   
            try
            {
                //Fajniejsza nazwa dla pliku raportu
                string OutputRaportFileName = "Raport_" + DateTime.Now.ToString("dd_MM_yyyy");
                //Czy debugujemy
                if (Interaction.ShowDebugDialog("Debuger", "Turn on debuger?"))
                    OurDebug.Enable();
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
                                OurDebug.AppendInfo("Email  ", DebugCorrectEmialsCounter.ToString(), ": ", email1.Subject, email1.ReceivedTime.ToString());
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
                    ExcelSheet raport = new ExcelSheet();

                    var row1 = 4;
                    var row2 = 4;
                    var row3 = 4;
                    emails = emailsWithoutDuplicates(emails);

                    foreach (MailItem newEmail in emails)
                    {
                        OurDebug.AppendInfo("Przed odczytem kategorii:", newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());//#endif
                        var typ = 0;
                        if (isMultipleCategoriesAndAnyOfTheireInterestedUs(newEmail.Categories))
                        {
                            OurDebug.AppendInfo("Po odczycie kategorii:", newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());
                            int emailConversationAmount = getConversationAmount(newEmail);
                            DateTime friday = getInflowDate();
                            typ = selectCorrectEmailType(newEmail);
                            OurDebug.AppendInfo("Nadany typ:", typ.ToString());
                            switch (typ)
                            {
                                case 1:
                                    row1++;
                                    raport.insertDataExcel(raport.oSheet, row1, newEmail, emailConversationAmount, 1);
                                    break;
                                case 2:
                                    row2++;
                                    raport.insertDataExcel(raport.oSheet, row2, newEmail, emailConversationAmount, 2);
                                    break;
                                case 3:
                                    row3++;
                                    raport.insertDataExcel(raport.oSheet, row3, newEmail, emailConversationAmount, 3);
                                    break;
                            }
                            raport.oSheet.Columns.AutoFit();
                            raport.oSheet.Cells[4, 1].EntireRow.Font.Bold = true;
                        }
                    }

                    raport.createCenterTables(raport.oSheet, row1, row2, row3);
                    raport.createExcelSumCategories(raport.oSheet, row1, row2, row3);
                    raport.oWB.SaveAs(OutputRaportFileName, Excel.XlFileFormat.xlOpenXMLStrictWorkbook);
                    raport.oWB.Close(true);
                    raport.oXL.Quit();
                    Marshal.ReleaseComObject(raport.oXL);
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
