﻿using System;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Diagnostics;

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
                Word.Application winword = new Word.Application();
                //Set animation status for word application  
                //winword.ShowAnimation = false;
                //Set status for word application is to be visible or not.  
                winword.Visible = false;
                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;
                //Create a new document  
                Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                /********************************************************************************************/


                string header = "NCMAILBOX tasks (week "; header += CurrentWeek(); header += "):";
                WriteMainHeader(header, document);
                string tym11 = "\tInflow: "; tym11 += Ribbon1.OurData.inflowAmount.ToString();
                WriteSecondHeader(tym11, document);
                WriteInflowMails(document);
                string tym22 = "\tIn-hands: "; tym22 += Ribbon1.OurData.inhandsAmount.ToString();
                WriteSecondHeader(tym22, document);
                WriteInhandsMails(document);
                string tym33 = "\tOutflow: "; tym33 += Ribbon1.OurData.outflowAmount.ToString();
                WriteSecondHeader(tym33, document);
                WriteOutflowMails(document);

                /**********************************************************************************************/
                //Save the document 
                object filename = path;
                document.SaveAs(ref filename, Word.WdSaveFormat.wdFormatDocumentDefault);
                document.Close(true);
                winword.Quit();
                Marshal.ReleaseComObject(winword);

            }
            catch (Exception ex)
            {
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Problem with createDocument Word. \n" , ex.StackTrace,"\n",ex.Message);
            }
        }
        void WriteMainHeader(string header, Word.Document document)
        {
            Word.Paragraph objPara;
            objPara = document.Paragraphs.Add();
            objPara.Range.Text = header;
            objPara.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            objPara.Range.Font.Size = 12;
            objPara.Range.Font.Italic = 0;
            objPara.Range.InsertParagraphAfter();
        }

        void WriteSecondHeader(string header, Word.Document document)
        {
            Word.Paragraph objPara;
            objPara = document.Paragraphs.Add();
            objPara.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            objPara.Range.Text = header;
            objPara.Range.Font.Size = 11;
            objPara.Range.Font.Italic = 0;
            objPara.Range.InsertParagraphAfter();
        }
        void WriteOutflowMails(Word.Document document)
        {
            foreach (string s in Ribbon1.OurData.outflow)
            {
                string tym = "\t\t"; tym += s;
                Word.Paragraph objPara;
                objPara = document.Paragraphs.Add();
                objPara.Range.Text = tym;
                objPara.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                objPara.Range.Font.Size = 10;
                objPara.Range.Font.Italic = 1;
                objPara.Range.InsertParagraphAfter();
            }
        }
        void WriteInflowMails(Word.Document document)
        {
            foreach (string s in Ribbon1.OurData.inflow)
            {
                string tym = "\t\t"; tym += s;
                Word.Paragraph objPara;
                objPara = document.Paragraphs.Add();
                objPara.Range.Text = tym;
                objPara.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                objPara.Range.Font.Size = 10;
                objPara.Range.Font.Italic = 1;
                objPara.Range.InsertParagraphAfter();
            }
        }
        void WriteInhandsMails(Word.Document document)
        {
            foreach (string s in Ribbon1.OurData.inhands)
            {
                string tym = "\t\t"; tym += s;
                Word.Paragraph objPara;
                objPara = document.Paragraphs.Add();
                objPara.Range.Text = tym;
                objPara.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                objPara.Range.Font.Size = 10;
                objPara.Range.Font.Italic = 1;
                objPara.Range.InsertParagraphAfter();
            }
        }
        int CurrentWeek()
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
