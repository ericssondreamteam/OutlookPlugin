using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;

namespace OutlookAddIn1
{
    class WordClass
    {
        public List<string> inflow = new List<string>();
        public List<string> outflow = new List<string>();
        public List<string> inhands = new List<string>();
        public int inflowAmount = 0;
        public int outflowAmount = 0;
        public int inhandsAmount = 0;
        private Hashtable myHashtable;

        public void addNewItem(string n, string k)
        {
            if (k == "inflow")
            {
                inflowAmount++;
                inflow.Add(n);
            }
            if (k == "outflow")
            {
                outflowAmount++;
                outflow.Add(n);
            }
            if (k == "inhands")
            {
                inhandsAmount++;
                inhands.Add(n);
            }
        }
        public void WriteToWord(string path)
        {
            CreateDocument(path);
        }
        public void CreateDocument(string path)
        {
            try
            {
                //Create an instance for word app  
                CheckWordProcesses();
                Word.Application winword = new Word.Application();
                int wordIDProcess = getWordID();

                //Set animation status for word application  
                winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                document.Content.Text += "NCMAILBOX tasks (week 24):";
                string tym1 = "\tInflow: "; tym1 += inflowAmount.ToString();
                document.Content.Text += tym1;
                foreach (string s in inflow)
                {
                    string tym = "\t\t"; tym += s; 
                    document.Content.Text += tym;
                    
                }

                string tym2 = "\tIn-hands: "; tym2 += inflowAmount.ToString();
                document.Content.Text += tym2;
                foreach (string s in inhands)
                {
                    string tym = "\t\t"; tym += s;
                    document.Content.Text += tym;
                }
                string tym3 = "\tOutflow: "; tym3 += outflowAmount.ToString();
                document.Content.Text += tym3;
                foreach (string s in outflow)
                {
                    string tym = "\t\t"; tym += s;
                    document.Content.Text += tym;
                }
                   
               



                //Save the document                 
                object filename = path;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                killWord(wordIDProcess);
                // MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                
            }
        }

        private void CheckWordProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("word");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process WordProcess in AllProcesses)
            {
                myHashtable.Add(WordProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        private int getWordID()
        {
            Process[] AllProcesses = Process.GetProcessesByName("word");
            foreach (Process WordProcess in AllProcesses)
            {
                if (myHashtable.ContainsKey(WordProcess.Id) == false)
                    return WordProcess.Id;
            }
            throw new SystemException("Process word.exe do not exist.");
        }

        /* Zabijamy proces ktory nie znajduje sie w hashtable */
        private void killWord(int processID)
        {
            Process[] AllProcesses = Process.GetProcessesByName("word");
            // check to kill the right process
            foreach (Process WordProcess in AllProcesses)
            {
                if (WordProcess.Id == processID)
                    WordProcess.Kill();
            }
        }
    }
}
