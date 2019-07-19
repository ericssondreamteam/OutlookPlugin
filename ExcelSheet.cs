using System;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Collections;

namespace OutlookAddIn1
{
    class ExcelSheet
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
            oSheet.Cells[3, 2] = "OUTFLOW";
            oSheet.Cells[3, 3] = "IN-HANDS";

            fillExcelCells(1, oSheet);
            fillExcelCells(2, oSheet);
            fillExcelCells(3, oSheet);
        }

        private void fillExcelCells(int i, Excel._Worksheet oSheet)
        {
            oSheet.Cells[4, i] = "Subject";
        }

        public void createExcelSumCategories(Excel._Worksheet oSheet, int rowInHands, int rowInflow, int rowOutflow)
        {
            oSheet.Cells[4, 5] = "SUMMARY";
            oSheet.Cells[5, 5] = "Inflow  = ";
            oSheet.Cells[6, 5] = "Outflow = ";
            oSheet.Cells[7, 5] = "In hands = ";

            if (rowInHands == 4) /* Gdy nie znajdzie zadnych maili w IN-HANDS */
                oSheet.Cells[7, 6].Value = 0;
            else
                oSheet.Cells[7, 6].Formula = "=ROWS(C5:C" + rowInHands + ")";
            if (rowInflow == 4) /* Gdy nie znajdzie zadnych maili w INFLOW */
                oSheet.Cells[5, 6].Value = 0;
            else
                oSheet.Cells[5, 6].Formula = "=ROWS(A5:A" + rowInflow + ")";
            if (rowOutflow == 4) /* Gdy nie znajdzie zadnych maili w OUTFLOW */
                oSheet.Cells[6, 6].Value = 0;
            else
                oSheet.Cells[6, 6].Formula = "=ROWS(B5:B" + rowOutflow + ")";
            oSheet.get_Range("E5", "E7").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        public void createCenterTables(Excel._Worksheet oSheet, int rowInHands, int rowInflow, int rowOutflow)
        {
            Excel.Range rangeForInflowTable = oSheet.get_Range("A4", "A" + rowInflow);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeForInflowTable,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "INFLOW";
            oSheet.ListObjects["INFLOW"].TableStyle = "TableStyleMedium9";

            Excel.Range rangeForOutflowTable = oSheet.get_Range("B4", "B" + rowOutflow);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeForOutflowTable,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "OUTFLOW";
            oSheet.ListObjects["OUTFLOW"].TableStyle = "TableStyleMedium12";

            Excel.Range rangeForInHandsTable = oSheet.get_Range("C4", "C" + rowInHands);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeForInHandsTable,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "IN-HANDS";
            oSheet.ListObjects["IN-HANDS"].TableStyle = "TableStyleMedium14";

            oSheet.get_Range("A5", "A" + rowInflow).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            oSheet.get_Range("B5", "B" + rowOutflow).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            oSheet.get_Range("C5", "C" + rowInHands).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
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
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (ExcelProcess.Id == processID)
                    ExcelProcess.Kill();
            }
        }

        public void SaveExcel(string OutputRaportFileName, Debuger ourDebug)
        {
            saveToExcel(OutputRaportFileName, ourDebug);
        }

        private void saveToExcel(string OutputRaportFileName, Debuger ourDebug)
        {
            try
            {
                var rowInHands = 4;
                var rowInflow = 4;
                var rowOutflow = 4;

                foreach (string s in Ribbon1.OurData.inflow)
                {
                    rowInflow++;
                    oSheet.Cells[rowInflow, 1] = s;
                }
                foreach (string s in Ribbon1.OurData.outflow)
                {
                    rowOutflow++;
                    oSheet.Cells[rowOutflow, 2] = s;
                }
                foreach (string s in Ribbon1.OurData.inhands)
                {
                    rowInHands++;
                    oSheet.Cells[rowInHands, 3] = s;
                }

                oSheet.Columns.AutoFit();
                oSheet.Cells[4, 1].EntireRow.Font.Bold = true;
                createCenterTables(oSheet, rowInHands, rowInflow, rowOutflow);
                createExcelSumCategories(oSheet, rowInHands, rowInflow, rowOutflow);
                oWB.SaveAs(OutputRaportFileName, Excel.XlFileFormat.xlOpenXMLStrictWorkbook);
                oWB.Close(true);
                oXL.Quit();
                killExcel(getExcelIDProcess());
            }
            catch(Exception ex)
            {
                ourDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Problem with saveToExcel function. \n" , ex.StackTrace,ex.Message);
            }
            
        }
    }
}
