using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using System.Runtime.InteropServices;
using System.Globalization;

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

                foreach (Word.Section section in document.Sections)
                {
                    //Get the header range and add the header details.  
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    //  headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    // headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                    headerRange.Font.Size = 12;
                   
                    string header = "NCMAILBOX tasks (week "; header += currentWeek(); header += "):";
                    headerRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    headerRange.Text = header;
                }
                //document.Content.Bold = 1;
                //document.Content.Font = "Calibri";
                // document.Content.Font.Size = 12;
                //document.Content.Italic = 1;
                document.Content.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                //document.Content.Font.Italic = 1;
                document.Content.Font.Name = "Calibri";
                document.Content.Text += "NCMAILBOX tasks (week 24):";
                document.Content.Font.Italic = 1;
                Word.Paragraph paraMain = document.Content.Paragraphs.Add(ref missing);
                object styleHeading1 = "Heading 1";
                paraMain.Range.set_Style(ref styleHeading1);
                //paraMain.Range.Text = "Para 1 text";
                paraMain.Range.InsertParagraphAfter();
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

        int currentWeek()
        {
            DateTime d = new DateTime();
            d = DateTime.Now;
            CultureInfo cul = CultureInfo.CurrentCulture;
            int weekNum = cul.Calendar.GetWeekOfYear(
                d,
                CalendarWeekRule.FirstDay,
                DayOfWeek.Monday);
            return weekNum;
        }


    }
}
