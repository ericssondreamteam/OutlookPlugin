using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using System.Diagnostics;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        static private Debuger OurDebug = new Debuger();
        private Office.IRibbonUI ribbon;
        static public DataObject OurData = new DataObject(OurDebug);
        WordClass toBeSavedWord = new WordClass();
        public static bool checkExcel = false;
        public static bool checkWord = false;
        private int DebugForEachCounter = 0;
        public static String fullInfoBox;

        public Ribbon1()
        {

        }



        public void OnTableButton(Office.IRibbonControl control)
        {
            Settings set = new Settings();
            try
            {
                string OutputRaportFileName = "Raport_" + DateTime.Now.ToString("dd_MM_yyyy");
                Form1 form3 = new Form1(ref OutputRaportFileName);
                form3.ShowDialog();
                if (Settings.ifWeDoRaport == DialogResult.OK)
                {
                    EmailFunctions functions = new EmailFunctions(OurDebug, Settings.boxMailName, DateTime.Parse(Settings.raportDate));

                    List<MailItem> emails = new List<MailItem>();
                    MailItem email1 = null;
                    int DebugCorrectEmailsCounter = 0;

                    functions.choiceOfFileFormat(Settings.checkList);


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
                    DebugForEachCounter = functions.getOnlyEmailsForTwoWeeksAgo(DebugForEachCounter, email1, oItems, DebugCorrectEmailsCounter, emails);

                    //Show how many times foreach is performed
                    OurDebug.AppendInfo("\n\n", "Ile razy foreach: ", DebugForEachCounter.ToString(), "Maile brane pod uwage po wstepnej selekcji: ", "\n\n");

                    //Delete duplicates from email in the same name or the same thread
                    try
                    {
                        emails = functions.emailsWithoutDuplicates(emails);
                        emails = functions.removeDuplicateOneMoreTime(emails);
                    }
                    catch (Exception e)
                    {
                        OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Usuwanie duplikatow nie dziala", e.StackTrace, "\n", e.Message);
                    }

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
                            OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Ribbon1.cs line:96. Problem in ID message.", ex.Message, "\n", ex.StackTrace);
                        }

                    }
                    OurData.lastTuning();
                    //Start create excel raport
                    if (checkExcel)
                    {
                        ExcelSheet raport = new ExcelSheet();
                        raport.SaveExcel(Settings.OutputRaportFileName, OurDebug);
                    }
                    //Save to txt file and word 
                    if (checkWord)
                    {
                        string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + Settings.OutputRaportFileName + ".docx";
                        toBeSavedWord.WriteToWord(path, OurDebug, DateTime.Parse(Settings.raportDate));
                    }

                    if (checkExcel)
                        fullInfoBox += "\n\nYour report (Excel) is saved: " + Settings.OutputRaportFileName + ".xlsx";
                    if (checkWord)
                        fullInfoBox += "\n\nYour report(Word) is saved: " + Settings.OutputRaportFileName + ".docx";


                    //Raport is saved
                    OurDebug.AppendInfo("Your report is SAVED :D");

                }
                else
                {
                    //OPERATION CANCELED
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Some error occured during second analysis\nIf You turn on debugger please go there");
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Ribbon1.cs line:135. SECOND TRY CATCH\n", e.Message, "\n", e.StackTrace);
            }
            finally
            {
                if (OurDebug.IsEnable())
                {
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    path += "\\DebugInfoRaportPlugin.txt";
                    OurDebug.SaveDebugInfoToFile(path);
                    fullInfoBox += "\n\nYour debug file is saved: DebugInfoRaportPlugin.txt";
                    OurDebug.Disable();
                }
                Form2 summary = new Form2();
                summary.ShowDialog();

                OurData.ClearData();
                DebugForEachCounter = 0;
                checkExcel = false;
                checkWord = false;
                fullInfoBox = String.Empty;
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
