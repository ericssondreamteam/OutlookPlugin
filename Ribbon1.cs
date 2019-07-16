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
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using System.Diagnostics;
using System.Collections;
using System.Threading;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {

        private Hashtable myHashtable;
        private Debuger OurDebug = new Debuger();
        public static int counter = 0;
        public static int progress = 0;
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
            today = today.AddDays(-2).AddHours(17);
            return today;
        }
        public int getConversationAmount(Outlook.MailItem newEmail)
        {
            try
            {
                Outlook.Conversation conv = newEmail.GetConversation();
                Outlook.Table table = conv.GetTable();
                return table.GetRowCount();
            }
            catch(Exception e)
            {
                OurDebug.AppendInfo("Blad w liczbie konwersacji");
                return 0;
            }

        }
        public int selectCorrectEmailType(Outlook.MailItem newEmail)
        {
            int typ = 0;
            if (newEmail.Categories == null) //inflow
            {
                if(newEmail.ReceivedTime > getInflowDate()) typ = 1;
                if (getConversationAmount(newEmail) > 1) typ = 1;
                else typ = 2;
            }
            if (getConversationAmount(newEmail) > 1 && newEmail.ReceivedTime > getInflowDate()) //in hands
            {
                typ = 1;
            }
            else if (getConversationAmount(newEmail) == 1 && newEmail.ReceivedTime > getInflowDate()) //inflow
            {
                typ = 2;
            }
            else if ((newEmail.ReceivedTime > getInflowDate().AddDays(-7)) && (newEmail.ReceivedTime < getInflowDate())) //outflow
            {
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

                        emails.RemoveAt(j);
                        j--;
                    }
                }
            }
            return emails;
        }
        public void OnTableButton(Office.IRibbonControl control)
        {   
            try
            {
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
                    Items oItems = oInbox2.Items;
                    List<MailItem> emails = new List<MailItem>();
                    var a = 2;
                    OurDebug.AppendInfo("Email's amount", oItems.Count.ToString());
                    oItems.Sort("[ReceivedTime]", true);//sortowanie od najnowszych wszystkich items 

                    MailItem email1 = null;
                    int DebugForEachCounter = 0;
                    int DebugCorrectEmailsCounter = 0;
                    OurDebug.AppendInfo("\n\n ************************MAILS*******************\n\n");
                    foreach (object collectionItem in oItems)
                    {
                        try
                        {
                            DebugForEachCounter++;
                            email1 = collectionItem as MailItem;
                            if (email1 != null)
                            {
                                OurDebug.AppendInfo("Email  ", DebugCorrectEmailsCounter.ToString(), ": ", email1.Subject, email1.ReceivedTime.ToString());
                                if (email1.ReceivedTime > getInflowDate().AddDays(-7))
                                {
                                    DebugCorrectEmailsCounter++;
                                    emails.Add(email1);
                                }
                                else
                                    break;
                            }
                            
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Some error occured during first analysis\nIf You turn on debugger please go there");
                            OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "FIRST TRY CATCH\n", "Emial number:", DebugCorrectEmailsCounter.ToString(), "\n", e.Message, "\n", e.StackTrace);
                        }
                    }
                    OurDebug.AppendInfo("\n\n", "Ile razy foreach: ", DebugForEachCounter.ToString(), "Maile brane pod uwage po wstepnej selekcji: ", "\n\n");

                    OurDebug.AppendInfo("\n\n", "Ile razy foreach: ", DebugForEachCounter.ToString(), "Maile brane pod uwage po wstepnej selekcji: ", "\n\n");
                    CheckExcellProcesses();
                    ExcelSheet raport = new ExcelSheet();
                    int processID = getExcelID();
                    var e2 = 2;
                    var rowInHands = 4;
                    var rowInflow = 4;
                    var rowOutflow = 4;
                    emails = emailsWithoutDuplicates(emails);

                    foreach (MailItem newEmail in emails)
                    {
                        progress++; 
                        Form1.incrementValue(progress);
                        OurDebug.AppendInfo("Przed odczytem kategorii:", newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());//#endif
                        var typ = 0;
                        if (isMultipleCategoriesAndAnyOfTheireInterestedUs(newEmail.Categories))
                        {
                            OurDebug.AppendInfo("Po odczycie kategorii:", newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());
                            int emailConversationAmount = getConversationAmount(newEmail); 
                            DateTime friday = getInflowDate();
                            typ = selectCorrectEmailType(newEmail);
                            OurDebug.AppendInfo("Nadany typ:", typ.ToString());
                            var a3 = 27;
                            switch (typ)
                            {
                                case 1:
                                    rowInHands++;
                                    raport.insertDataExcel(raport.oSheet, rowInHands, newEmail, emailConversationAmount, 1);
                                    break;
                                case 2:
                                    rowInflow++;
                                    raport.insertDataExcel(raport.oSheet, rowInflow, newEmail, emailConversationAmount, 2);
                                    break;
                                case 3:
                                    rowOutflow++;
                                    raport.insertDataExcel(raport.oSheet, rowOutflow, newEmail, emailConversationAmount, 3);
                                    break;
                            }
                            raport.oSheet.Columns.AutoFit();
                            raport.oSheet.Cells[4, 1].EntireRow.Font.Bold = true;
                        }
                    }

                    raport.createCenterTables(raport.oSheet, rowInHands, rowInflow, rowOutflow);
                    raport.createExcelSumCategories(raport.oSheet, rowInHands, rowInflow, rowOutflow);
                    raport.oWB.SaveAs(OutputRaportFileName, Excel.XlFileFormat.xlOpenXMLStrictWorkbook);
                    raport.oWB.Close(true);
                    raport.oXL.Quit();
                    KillExcel(processID);
                    //Marshal.ReleaseComObject(raport.oXL);
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
            OurDebug.AppendInfo("Categories start:",categories);
            if (categories is null)
            {
                return true;
            }
            else
            {
                var a = 0;
                categories = categories.Trim();
                categories = categories.Replace(" ", "");
                categories = categories.ToLower();
                OurDebug.AppendInfo("Categories after trim and repalce and lower:", categories);
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
        /* zapisujemy id procesow do hashTable przed uruchomieniem naszego procesu */
        private void CheckExcellProcesses() 
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        private int getExcelID()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            foreach(Process ExcelProcess in AllProcesses)
            {
                if (myHashtable.ContainsKey(ExcelProcess.Id) == false)
                    return ExcelProcess.Id;
            }
            throw new SystemException("Process excel.exe do not exist. Check constructor in class 'ExcelSheet'");
        }

        /* Zabijamy proces ktory nie znajduje sie w hashtable */
        private void KillExcel(int processID)
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            // check to kill the right process
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (ExcelProcess.Id == processID)
                    ExcelProcess.Kill();
            }
        }
    }

}
