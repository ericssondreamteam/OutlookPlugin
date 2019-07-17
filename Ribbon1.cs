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

        public Ribbon1()
        {

        }

        public void OnTableButton(Office.IRibbonControl control)
        {   
            try
            {
                //Initialize
                EmailFunctions functions = new EmailFunctions(OurDebug);
                string OutputRaportFileName = "Raport_" + DateTime.Now.ToString("dd_MM_yyyy");
                List<MailItem> emails = new List<MailItem>();
                MailItem email1 = null;
                int DebugForEachCounter = 0;
                int DebugCorrectEmailsCounter = 0;

                //Czy debugujemy
                if (Interaction.ShowDebugDialog("Debuger", "Turn on debuger?"))
                {
                    OurDebug.Enable();
                }
                    
                else
                {
                    OurDebug.Disable();
                }

                if (Interaction.SaveRaportDialog("New document", "New document name:", ref OutputRaportFileName) == DialogResult.OK)
                {
                    //Initialize outlook app
                    Outlook.Application oApp = new Outlook.Application();
                    NameSpace oNS = oApp.GetNamespace("mapi");
                    MAPIFolder oInbox2 = oApp.ActiveExplorer().CurrentFolder as MAPIFolder;
                    OurDebug.AppendInfo("Wybrany folder ", oInbox2.Name);
                    Items oItems = oInbox2.Items;
                    OurDebug.AppendInfo("Email's amount", oItems.Count.ToString());

                    //Sort all items
                    oItems.Sort("[ReceivedTime]", true);

                    //Debug info for mails
                    OurDebug.AppendInfo("\n\n ************************MAILS*******************\n\n");

                    //Get only mails from two weeks ago
                    functions.getOnlyEmailsForTwoWeeksAgo(DebugForEachCounter, email1, oItems, DebugCorrectEmailsCounter, emails);

                    //Show how many times foreach is performed
                    OurDebug.AppendInfo("\n\n", "Ile razy foreach: ", DebugForEachCounter.ToString(), "Maile brane pod uwage po wstepnej selekcji: ", "\n\n");


                    CheckExcellProcesses();
                    ExcelSheet raport = new ExcelSheet();
                    int processID = getExcelID();
                    var rowInHands = 4;
                    var rowInflow = 4;
                    var rowOutflow = 4;
                    emails = functions.emailsWithoutDuplicates(emails);
                    emails = functions.removeDuplicateOneMoreTime(emails);
                    foreach (MailItem newEmail in emails)
                    {
                        OurDebug.AppendInfo("Przed odczytem kategorii:", newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());//#endif
                        var typ = 0;
                        if (functions.isMultipleCategoriesAndAnyOfTheireInterestedUs(newEmail.Categories))
                        {
                            OurDebug.AppendInfo("Po odczycie kategorii:", newEmail.Subject, newEmail.Categories, newEmail.ReceivedTime.ToString());
                            int emailConversationAmount = functions.getConversationAmount(newEmail); 
                            DateTime friday = functions.getInflowDate();
                            typ = functions.selectCorrectEmailType(newEmail);
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
                    WriteToTxtFile(WriteInCorrextFomrat(koncowaLista), path);
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
            koncowyString.Append("Inflow: "+tematy.inflowAmount+"\n");
            int i;
            for (i = 0; i < tematy.inflow.Count; i++)
                koncowyString.Append("\t" + tematy.inflow[i] + "\n");
            koncowyString.Append("In-hands: " + tematy.inhandsAmount + "\n");
            for (i = 0; i < tematy.inhands.Count; i++)
                koncowyString.Append("\t" + tematy.inhands[i] + "\n");
            koncowyString.Append("Outflow: " + tematy.outflowAmount + "\n");
            for (i = 0; i < tematy.outflow.Count; i++)
                koncowyString.Append("\t" + tematy.outflow[i] + "\n");

            return koncowyString;

        }
        private void WriteToTxtFile(StringBuilder doZapisu,string path)
        {
            MessageBox.Show(path);
            File.WriteAllText(path, doZapisu.ToString());           
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
