using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;

namespace OutlookAddIn1
{
    class WordClass
    {
        private Hashtable myHashtable;
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

                //Add header into the document  
                foreach (Word.Section section in document.Sections)
                {
                    //Get the header range and add the header details.  
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    headerRange.Text = "Header text goes here";
                }

                //Add the footers into the document  
                foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                {
                    //Get the footer range and add the footer details.  
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = "Footer text goes here";
                }

                //adding text to document  
                document.Content.SetRange(0, 0);
                //  document.Content.Text = "This is test document " + Environment.NewLine;

                //Add paragraph with Heading 1 style  
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                object styleHeading1 = "Heading 1";
                para1.Range.set_Style(ref styleHeading1);
                para1.Range.Text = "Para 1 text";
                para1.Range.InsertParagraphAfter();

                //Add paragraph with Heading 2 style  
                Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
                object styleHeading2 = "Heading 2";
                para2.Range.set_Style(ref styleHeading2);
                para2.Range.Text = "Para 2 text";
                para2.Range.InsertParagraphAfter();



                //Save the document  
                MessageBox.Show("TEST "+path);
                object filename = path;
                MessageBox.Show("CO TO: "+filename.ToString());
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
                // MessageBox.Show(ex.Message);
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
