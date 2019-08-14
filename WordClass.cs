using System;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace OutlookAddIn1
{
    class WordClass
    {
        public void WriteToWord(string path, Debuger OurDebug)
        {
            CreateDocument(path, OurDebug);
        }
        private void CreateDocument(string path, Debuger OurDebug)
        {
            try
            {
                /********************************************************************************************/
                //Create an instance for word app              
                Application winword = new Application();
                //Set animation status for word application  
                //winword.ShowAnimation = false;
                //Set status for word application is to be visible or not.  
                winword.Visible = false;
                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;
                //Create a new document  
                Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                /********************************************************************************************/


                string header = "NCMAILBOX tasks (week "; header += CurrentWeek(); header += "):";
                WriteMainHeader(header, document);
                string tym11 = "\tInflow: "; tym11 += Ribbon1.OurData.inflowAmount.ToString();
                WriteSecondHeader(tym11, document);
                WriteMails(document, Ribbon1.OurData.inflow);
                string tym22 = "\tIn-hands: "; tym22 += Ribbon1.OurData.inhandsAmount.ToString();
                WriteSecondHeader(tym22, document);
                WriteMails(document, Ribbon1.OurData.inhands);
                string tym33 = "\tOutflow: "; tym33 += Ribbon1.OurData.outflowAmount.ToString();
                WriteSecondHeader(tym33, document);
                WriteMails(document, Ribbon1.OurData.outflow);

                /**********************************************************************************************/
                //Save the document 
                object filename = path;
                document.SaveAs(ref filename, WdSaveFormat.wdFormatDocumentDefault);
                document.Close(true);
                winword.Quit();
                Marshal.ReleaseComObject(winword);

            }
            catch (Exception ex)
            {
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Problem with createDocument Word. \n", ex.StackTrace, "\n", ex.Message);
            }
        }
        private void WriteMainHeader(string header, Document document)
        {
            Paragraph objPara;
            objPara = document.Paragraphs.Add();
            objPara.Range.Text = header;
            objPara.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
            objPara.Range.Font.Size = 12;
            objPara.Range.Font.Italic = 0;
            objPara.Range.InsertParagraphAfter();
        }

        private void WriteSecondHeader(string header, Document document)
        {
            Paragraph objPara;
            objPara = document.Paragraphs.Add();
            objPara.Range.Font.Underline = WdUnderline.wdUnderlineNone;
            objPara.Range.Text = header;
            objPara.Range.Font.Size = 11;
            objPara.Range.Font.Italic = 0;
            objPara.Range.InsertParagraphAfter();
        }

        private void WriteMails(Document document, List<string> list)
        {
            foreach (string s in list)
            {
                string tym = "\t\t"; tym += s;
                Paragraph objPara;
                objPara = document.Paragraphs.Add();
                objPara.Range.Text = tym;
                objPara.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                objPara.Range.Font.Size = 10;
                objPara.Range.Font.Italic = 1;
                objPara.Range.InsertParagraphAfter();
            }
        }

        private int CurrentWeek()
        {
            DateTime d = DateTime.Now;
            CultureInfo cul = CultureInfo.CurrentCulture;
            int weekNum = cul.Calendar.GetWeekOfYear(
                d,
                CalendarWeekRule.FirstDay,
                DayOfWeek.Monday);
            return weekNum;
        }

    }
}