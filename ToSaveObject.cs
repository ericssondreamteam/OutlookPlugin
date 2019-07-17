using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Word=Microsoft.Office.Interop.Word;



namespace OutlookAddIn1
{
    class ToSaveObject
    {
        private List<string> inflow = new List<string>();
        private List<string> outflow = new List<string>();
        private List<string> inhands = new List<string>();
        private int inflowAmount = 0;
        private int outflowAmount = 0;
        private int inhandsAmount = 0;

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
        private StringBuilder WriteInCorrextFomrat()
        {
            StringBuilder koncowyString = new StringBuilder();
            int i;
            koncowyString.Append("Inflow: " + inflowAmount + "\n");
            for (i = 0; i < inflow.Count; i++)
                koncowyString.Append("\t" + inflow[i] + "\n");
            koncowyString.Append("In-hands: " + inhandsAmount + "\n");
            for (i = 0; i < inhands.Count; i++)
                koncowyString.Append("\t" + inhands[i] + "\n");
            koncowyString.Append("Outflow: " + outflowAmount + "\n");
            for (i = 0; i < outflow.Count; i++)
                koncowyString.Append("\t" + outflow[i] + "\n");

            return koncowyString;

        }
        public void WriteToTxtFile(string path)
        {
            File.WriteAllText(path, WriteInCorrextFomrat().ToString());
            inhands.Clear();
            inflow.Clear();
            outflow.Clear();
            inflowAmount = 0;
            outflowAmount = 0;
            inhandsAmount = 0;
        }
        public void WriteToWord(string path)
        {
            CreateDocument();           
        }
        private void CreateDocument()
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
                object filename = @"c:\temp1.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                // MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
        }
    }
}
