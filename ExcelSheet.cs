using System;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Collections;

namespace OutlookAddIn1
{
    public class ExcelSheet
    {
        private Hashtable myHashtable;
        public Excel.Application oXL;
        public Excel._Workbook oWB;
        public Excel._Worksheet oSheet;
        private int excelIDProcess;
        public ExcelSheet()
        {
            checkExcellProcesses();
            oXL = new Excel.Application();
            oXL.Visible = false;
            oWB = (oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            createExcelColumn(oSheet);
            excelIDProcess = getExcelID();
        }

        public int getExcelIDProcess()
        {
            return excelIDProcess;
        }
        public void createExcelColumn(Excel._Worksheet oSheet)
        {
            oSheet.Cells[1, 1] = "Raport Time: " + DateTime.Now.ToLongTimeString();
            oSheet.Cells[1, 2] = "Raport Date: " + DateTime.Now.ToLongDateString();
            oSheet.Cells[3, 1] = "INFLOW";
            oSheet.Cells[3, 5] = "OUTFLOW";
            oSheet.Cells[3, 9] = "IN-HANDS";

            fillExcelCells(1, oSheet);
            fillExcelCells(5, oSheet);
            fillExcelCells(9, oSheet);
        }

        private void fillExcelCells(int i, Excel._Worksheet oSheet)
        {
            oSheet.Cells[4, i] = "Subject";
            oSheet.Cells[4, i + 1] = "Messages amount";
            oSheet.Cells[4, i + 2] = "Category";
        }

        public void createExcelSumCategories(Excel._Worksheet oSheet, int rowInHands, int rowInflow, int rowOutflow)
        {
            oSheet.Cells[4, 13] = "SUMMARY";
            oSheet.Cells[5, 13] = "Inflow  = ";
            oSheet.Cells[6, 13] = "Outflow = ";
            oSheet.Cells[7, 13] = "In hands = ";

            if (rowInHands == 4) /* Gdy nie znajdzie zadnych maili w IN-HANDS */
                oSheet.Cells[7, 14].Value = 0;
            else
                oSheet.Cells[7, 14].Formula = "=ROWS(I5:I" + rowInHands + ")";
            if (rowInflow == 4) /* Gdy nie znajdzie zadnych maili w INFLOW */
                oSheet.Cells[5, 14].Value = 0;
            else
                oSheet.Cells[5, 14].Formula = "=ROWS(A5:A" + rowInflow + ")";
            if (rowOutflow == 4) /* Gdy nie znajdzie zadnych maili w OUTFLOW */
                oSheet.Cells[6, 14].Value = 0;
            else
                oSheet.Cells[6, 14].Formula = "=ROWS(E5:E" + rowOutflow + ")";
            oSheet.get_Range("N5", "N7").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        public void insertDataExcel(Excel._Worksheet oSheet, int row, Outlook.MailItem newEmail, int amount, int whichCategory)
        {
            if (whichCategory == 1) //IN-HANDS
                setDataIntoCells(oSheet, row, newEmail, amount, 9);
            else if (whichCategory == 2) //INFLOW
                setDataIntoCells(oSheet, row, newEmail, amount, 1);
            else if (whichCategory == 3) //OUTFLOW
                setDataIntoCells(oSheet, row, newEmail, amount, 5);
        }

        public void insertDataExcelInflowInHands(Excel._Worksheet oSheet, int rowInflow, int rowInHands, Outlook.MailItem newEmail, int amount)
        {
            setDataIntoCells(oSheet, rowInflow, newEmail, amount, 1);
            setDataIntoCells(oSheet, rowInHands, newEmail, amount, 9);
        }

        private void setDataIntoCells(Excel._Worksheet oSheet, int row, Outlook.MailItem newEmail, int amount, int whichColumn)
        {
            //if (newEmail.Subject.Substring(0, 3).Equals("RE:") || newEmail.Subject.Substring(0, 3).Equals("FW:"))
            //    oSheet.Cells[row, whichColumn] = newEmail.Subject.Substring(4);
            //else
                oSheet.Cells[row, whichColumn] = newEmail.Subject;
            oSheet.Cells[row, whichColumn + 1] = amount;
            oSheet.Cells[row, whichColumn + 2] = newEmail.Categories;
          //  oSheet.Cells[row, whichColumn + 3] = newEmail.ReceivedTime;
        }

        public void createCenterTables(Excel._Worksheet oSheet, int rowInHands, int rowInflow, int rowOutflow)
        {
            Excel.Range rangeForInflowTable = oSheet.get_Range("A4", "C" + rowInflow);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeForInflowTable,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "INFLOW";
            oSheet.ListObjects["INFLOW"].TableStyle = "TableStyleMedium9";

            Excel.Range rangeForOutflowTable = oSheet.get_Range("E4", "G" + rowOutflow);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeForOutflowTable,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "OUTFLOW";
            oSheet.ListObjects["OUTFLOW"].TableStyle = "TableStyleMedium12";

            Excel.Range rangeForInHandsTable = oSheet.get_Range("I4", "K" + rowInHands);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeForInHandsTable,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "IN-HANDS";
            oSheet.ListObjects["IN-HANDS"].TableStyle = "TableStyleMedium14";

            oSheet.get_Range("B5", "B" + rowInflow).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            oSheet.get_Range("F5", "F" + rowOutflow).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            oSheet.get_Range("J5", "J" + rowInHands).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        /* zapisujemy id procesow do hashTable przed uruchomieniem naszego procesu */
        private void checkExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        private int getExcelID()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (myHashtable.ContainsKey(ExcelProcess.Id) == false)
                    return ExcelProcess.Id;
            }
            throw new SystemException("Process excel.exe do not exist. Check constructor in class 'ExcelSheet'");
        }

        /* Zabijamy proces ktory nie znajduje sie w hashtable */
        public void killExcel(int processID)
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            // check to kill the right process
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (ExcelProcess.Id == processID)
                    ExcelProcess.Kill();
            }
        }
    }
}
