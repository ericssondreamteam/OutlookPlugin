using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace OutlookAddIn1
{
    public class ExcelSheet
    {
        public Excel.Application oXL;
        public Excel._Workbook oWB;
        public Excel._Worksheet oSheet;
        public ExcelSheet()
        {
            oXL = new Excel.Application();
            oXL.Visible = false;
            oWB = (oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            createExcelColumn(oSheet);
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

        public void createExcelSumCategories(Excel._Worksheet oSheet, int row1, int row2, int row3)
        {
            oSheet.Cells[4, 13] = "SUMMARY";
            oSheet.Cells[5, 13] = "Inflow  = ";
            oSheet.Cells[6, 13] = "Outflow = ";
            oSheet.Cells[7, 13] = "In hands = ";

            if (row1 == 4) /* Gdy nie znajdzie zadnych maili w IN-HANDS */
                oSheet.Cells[7, 14].Value = 0;
            else
                oSheet.Cells[7, 14].Formula = "=ROWS(I5:F" + row1 + ")";
            if (row2 == 4) /* Gdy nie znajdzie zadnych maili w INFLOW */
                oSheet.Cells[5, 14].Value = 0;
            else
                oSheet.Cells[5, 14].Formula = "=ROWS(A5:A" + row2 + ")";
            if (row3 == 4) /* Gdy nie znajdzie zadnych maili w OUTFLOW */
                oSheet.Cells[6, 14].Value = 0;
            else
                oSheet.Cells[6, 14].Formula = "=ROWS(E5:E" + row3 + ")";
            oSheet.get_Range("N5", "N7").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        }

        public void insertDataExcel(Excel._Worksheet oSheet, int row, Outlook.MailItem newEmail, int amount, int whichCategory)
        {
            if (whichCategory == 1) //IN-HANDS
            {
                oSheet.Cells[row, 9] = newEmail.Subject;
                oSheet.Cells[row, 10] = amount;
                oSheet.Cells[row, 11] = newEmail.Categories;
                oSheet.Cells[row, 12] = newEmail.ConversationID;
            }
            if (whichCategory == 2) //INFLOW
            {
                oSheet.Cells[row, 1] = newEmail.Subject;
                oSheet.Cells[row, 2] = amount;
                oSheet.Cells[row, 3] = newEmail.Categories;
            }
            if (whichCategory == 3) //OUTFLOW
            {
                oSheet.Cells[row, 5] = newEmail.Subject;
                oSheet.Cells[row, 6] = amount;
                oSheet.Cells[row, 7] = newEmail.Categories;
            }

        }
        public void createCenterTables(Excel._Worksheet oSheet, int row1, int row2, int row3)
        {
            Excel.Range tRange1 = oSheet.get_Range("A4", "C" + row2);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange1,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "INFLOW";
            oSheet.ListObjects["INFLOW"].TableStyle = "TableStyleMedium9";

            Excel.Range tRange2 = oSheet.get_Range("E4", "G" + row3);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange2,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "OUTFLOW";
            oSheet.ListObjects["OUTFLOW"].TableStyle = "TableStyleMedium12";

            Excel.Range tRange3 = oSheet.get_Range("I4", "K" + row1);
            oSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange3,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "IN-HANDS";
            oSheet.ListObjects["IN-HANDS"].TableStyle = "TableStyleMedium14";

            oSheet.get_Range("B5", "B" + row2).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            oSheet.get_Range("F5", "F" + row3).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            oSheet.get_Range("J5", "J" + row1).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }
    }
}
