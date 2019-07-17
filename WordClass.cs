using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using System.Runtime.InteropServices;

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
                Word.Application winword = new Word.Application();             

                //Set animation status for word application  
                winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                //document.Content.Bold = 1;
                //document.Content.Font = "Calibri";
               // document.Content.Font.Size = 12;
                //document.Content.Italic = 1;
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
                Marshal.ReleaseComObject(winword);

            }
            catch (Exception ex)
            {
                
            }
        }

      
    }
}
