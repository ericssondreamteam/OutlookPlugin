using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace OutlookAddIn1
{
    public partial class Loading : Form
    {

        static public Debuger OurDebug = new Debuger();
        //private Office.IRibbonUI ribbon;
        static public DataObject OurData = new DataObject(OurDebug);
        WordClass toBeSavedWord = new WordClass();
        public static bool checkExcel = false;
        public static bool checkWord = false;
        public static int DebugForEachCounter = 0;
        public static String fullInfoBox;
        public Loading()
        {
            InitializeComponent();
            backgroundWorker1.RunWorkerAsync();
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void ProgressBar1_Click(object sender, EventArgs e)
        {

        }

        private void pb_DoWork(object sender, DoWorkEventArgs e)
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
            catch (Exception ex)
            {
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Usuwanie duplikatow nie dziala", ex.StackTrace, "\n", ex.Message);
            }

            int counterForAllEmails = 0;
           
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
                counterForAllEmails++;
                backgroundWorker1.ReportProgress(counterForAllEmails / emails.Count * 60);
            }
            OurData.lastTuning();
            //Start create excel raport
            if (checkExcel)
            {
                //textBox1.Text = "zapis excela";
                ExcelSheet raport = new ExcelSheet();
                raport.SaveExcel(Settings.OutputRaportFileName, OurDebug);
            }
            backgroundWorker1.ReportProgress(80);
            //Save to txt file and word 
            if (checkWord)
            {
                //textBox1.Text = "zapis worda";
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + Settings.OutputRaportFileName + ".docx";
                toBeSavedWord.WriteToWord(path, OurDebug, DateTime.Parse(Settings.raportDate));
            }
            backgroundWorker1.ReportProgress(100);
            Thread.Sleep(1000);

            if (checkExcel)
                fullInfoBox += "\n\nYour report (Excel) is saved: " + Settings.OutputRaportFileName + ".xlsx";
            if (checkWord)
                fullInfoBox += "\n\nYour report(Word) is saved: " + Settings.OutputRaportFileName + ".docx";


            //Raport is saved
            OurDebug.AppendInfo("Your report is SAVED :D");


        }

        private void pb_Progress(object sender, ProgressChangedEventArgs e)
        {
         
            if (progressBar1.Value > 40 && progressBar1.Value <= 60 && checkExcel)
                label2.Text = "Excel is being created";
            else if(progressBar1.Value > 60 && checkWord)
                label2.Text = "Word is being created";
            progressBar1.Value = e.ProgressPercentage;          

        }

        private void pb_Done(object sender, RunWorkerCompletedEventArgs e)
        {
            Close();
        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Label2_Click_1(object sender, EventArgs e)
        {

        }
    }
}
