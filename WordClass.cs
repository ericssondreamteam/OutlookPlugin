using System;
using System.Collections.Generic;

using Word = Microsoft.Office.Interop.Word;

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
                   
                ////Add header into the document  
                //foreach (Word.Section section in document.Sections)
                //{
                //    //Get the header range and add the header details.  
                //    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                //    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //    headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                //    headerRange.Font.Size = 10;
                //    headerRange.Text = "Header text goes here";
                //}

                ////Add the footers into the document  
                //foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                //{
                //    //Get the footer range and add the footer details.  
                //    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                //    footerRange.Font.Size = 10;
                //    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //    footerRange.Text = "Footer text goes here";
                //}

                ////adding text to document  
                //document.Content.SetRange(0, 0);
                ////  document.Content.Text = "This is test document " + Environment.NewLine;

                ////Add paragraph with Heading 1 style  
                //Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                //object styleHeading1 = "Heading 1";
                //para1.Range.set_Style(ref styleHeading1);
                //para1.Range.Text = "Para 1 text";
                //para1.Range.InsertParagraphAfter();

                ////Add paragraph with Heading 2 style  
                //Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
                //object styleHeading2 = "Heading 2";
                //para2.Range.set_Style(ref styleHeading2);
                //para2.Range.Text = "Para 2 text";
                //para2.Range.InsertParagraphAfter();



                //Save the document                 
                object filename = path;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            }
            catch (Exception ex)
            {
                
            }
        }
    }
}
