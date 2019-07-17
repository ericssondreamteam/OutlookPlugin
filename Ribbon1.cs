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
using System.Text;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {

        private Hashtable myHashtable;
        private Debuger OurDebug = new Debuger();
        private Office.IRibbonUI ribbon;
        ToSaveObject koncowaLista = new ToSaveObject();


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
        public int getConversationAmount(MailItem newEmail)
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
        public int selectCorrectEmailType(MailItem newEmail)
        {
            int typ = 0;
            if(newEmail.Categories != null)
            {
                if (getConversationAmount(newEmail) > 1 && newEmail.ReceivedTime > getInflowDate()) //in hands
                {
                    typ = 1;
                }
                else if (newEmail.ReceivedTime > getInflowDate()) //inflow
                {
                    typ = 2;
                }
                else if ((newEmail.ReceivedTime > getInflowDate().AddDays(-7)) && (newEmail.ReceivedTime < getInflowDate())) //outflow
                {
                    typ = 3;
                }
                if (typ == 1) //inflow + in hands
                {
                    typ = 4;
                }
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
                    var rowInHands = 4;
                    var rowInflow = 4;
                    var rowOutflow = 4;
                    emails = emailsWithoutDuplicates(emails);
                    emails = removeDuplicateOneMoreTime(emails);
                   // int ktoryElemntZListyElementowXD = 0;
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
                                    rowInHands++;
                                    raport.insertDataExcel(raport.oSheet, rowInHands, newEmail, emailConversationAmount, 1);
                                    koncowaLista.addNewItem(newEmail.Subject,"inhands");
                                    break;
                                case 2:
                                    rowInflow++;
                                    raport.insertDataExcel(raport.oSheet, rowInflow, newEmail, emailConversationAmount, 2);
                                    koncowaLista.addNewItem(newEmail.Subject, "inflow");
                                    break;
                                case 3:
                                    rowOutflow++;
                                    raport.insertDataExcel(raport.oSheet, rowOutflow, newEmail, emailConversationAmount, 3);
                                    koncowaLista.addNewItem(newEmail.Subject, "outflow");
                                    break;
                                case 4:
                                    rowInflow++;
                                    rowInHands++;
                                    raport.insertDataExcelInflowInHands(raport.oSheet, rowInflow, rowInHands, newEmail, emailConversationAmount);
                                    koncowaLista.addNewItem(newEmail.Subject, "inhands");
                                    koncowaLista.addNewItem(newEmail.Subject, "inflow");
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
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    path += "\\";
                    path += OutputRaportFileName;
                    path += ".txt";
                    WriteToTxtFile(WriteInCorrextFomrat(koncowaLista),path);

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
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    path += "\\DebugInfoRaportPlugin.txt";
                    OurDebug.SaveDebugInfoToFile(path);
                    MessageBox.Show("Plik debugowania zapisany w "+path);
                }
            }
        }
        private StringBuilder WriteInCorrextFomrat(ToSaveObject tematy)
        {
            StringBuilder koncowyString = new StringBuilder();
            koncowyString.Append("Inflow "+tematy.inflowAmount+"\n");
            int i;
            for (i = 0; i < tematy.inflow.Count; i++)
                koncowyString.Append("\t" + tematy.inflow[i] + "\n");
            koncowyString.Append("In-hands " + tematy.inflowAmount + "\n");
            for (i = 0; i < tematy.inhands.Count; i++)
                koncowyString.Append("\t" + tematy.inhands[i] + "\n");
            koncowyString.Append("Outflow " + tematy.outflowAmount + "\n");
            for (i = 0; i < tematy.outflow.Count; i++)
                koncowyString.Append("\t" + tematy.outflow[i] + "\n");

            return koncowyString;

        }
        private void WriteToTxtFile(StringBuilder doZapisu,string path)
        {
            MessageBox.Show(path);
            File.WriteAllText(path, doZapisu.ToString());           
        }

        private List<MailItem> removeDuplicateOneMoreTime(List<MailItem> emails)
        {
            string mailSubject1;
            string mailSubject2;
            for (int i = 0; i < emails.Count - 1; i++)
            {
                mailSubject1 = emails[i].Subject;
                mailSubject1 = mailSubject1.Trim();
                mailSubject1 = mailSubject1.Replace(" ", "");
                mailSubject1 = mailSubject1.ToLower();
                if (mailSubject1.Substring(0, 3).Equals("re:") || mailSubject1.Substring(0, 3).Equals("fw:"))
                    mailSubject1 = mailSubject1.Substring(3);
              
                for (int j = i + 1; j < emails.Count; j++)
                {
                    mailSubject2 = emails[j].Subject;
                    mailSubject2 = mailSubject2.Trim();
                    mailSubject2 = mailSubject2.Replace(" ", "");
                    mailSubject2 = mailSubject2.ToLower();
                    if (mailSubject2.Substring(0, 3).Equals("re:") || mailSubject2.Substring(0, 3).Equals("fw:"))
                        mailSubject2 = mailSubject2.Substring(3);

                    if (mailSubject1.Equals(mailSubject2))

                    {
  
                        emails.RemoveAt(j);
                        j--;
                    }
                }
            }
            return emails;
        }


        bool isMultipleCategoriesAndAnyOfTheireInterestedUs(string categories)
        {
            OurDebug.AppendInfo("Categories start:",categories);
            if (categories is null)
            {
                return false;
            }
            else
            {
                categories = categories.Trim();
                categories = categories.Replace(" ", "");
                categories = categories.ToLower();
                OurDebug.AppendInfo("Categories after trim and repalce and lower:", categories);
                string[] categoriesList = categories.Split(',');
                foreach (var cat in categoriesList)
                {   //No Response Necessary    or    Unknown     No Response Necessary, Unknown
                    if (!cat.Equals("noresponsenecessary") && !cat.Equals("unknown") && !cat.Equals(""))
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
