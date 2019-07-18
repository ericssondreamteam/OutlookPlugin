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
        private Debuger OurDebug = new Debuger();
        private Office.IRibbonUI ribbon;
        ToSaveObject endingCorrectList = new ToSaveObject();
        ToSaveObject toBeSavedTemp = new ToSaveObject();
        ToSaveObject toBeSavedTemp1 = new ToSaveObject();
        WordClass toBeSavedWord = new WordClass();
        static public DataObject OurData = new DataObject();

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

                List<bool> checkList = Interaction.ShowDebugDialog("Debuger", "Excel", "Txt", "CheckBoxes");
                if (checkList[0])
                    OurDebug.Enable();
                //if(checkList[1])
                //Enable Excel
                //if(checkList[2])
                //Enable Txt

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

                    //Delete duplicates from email in the same name or the same thread
                    emails = functions.emailsWithoutDuplicates(emails);
                    emails = functions.removeDuplicateOneMoreTime(emails);

                    //Iterate all emails
                    foreach (MailItem newEmail in emails)
                    {
                        try
                        {
                            List<bool> categoryList;
                            //Divide on category
                            if (functions.isMultipleCategoriesAndAnyOfTheireInterestedUs(newEmail.Categories))
                            {
                                //Get inflow date and set to category
                                DateTime friday = functions.getInflowDate();
                                categoryList = functions.selectCorrectEmailType(newEmail);
                                OurData.addNewItem(newEmail.Subject, categoryList);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Nasz try catch vol.3 - Problem w ID wiadomosci");
                        }
                        
                    }

                    //Start create excel raport
                    ExcelSheet raport = new ExcelSheet();
                    raport.saveToExcel(OutputRaportFileName);

                    //Save to txt file and word
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + OutputRaportFileName + ".docx";
                    endingCorrectList.WriteToTxtFile(path);
                    toBeSavedWord.WriteToWord(path);
                    OurData.ClearData();

                    //Raport is saved
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
                    MessageBox.Show("Plik debugowania zapisany w " + path);
                }
            }
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
