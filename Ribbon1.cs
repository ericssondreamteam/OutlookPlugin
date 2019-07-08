
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

namespace OutlookAddIn1
{


    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {

        private Office.IRibbonUI ribbon;
        public void OnTextButton(Office.IRibbonControl control)
        {
            MessageBox.Show("You clicked a different control."+control.Id);
        }
        public void OnTableButton(Office.IRibbonControl control)
        {
            Outlook.Application oApp = new Outlook.Application();
            Outlook.NameSpace oNS = oApp.GetNamespace("mapi");
            Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items oItems = oInbox.Items;
            /*Outlook.MailItem oMsg = (Outlook.MailItem)oItems.GetFirst();
            String a = oMsg.ReceivedTime + " " + oMsg.SenderName;
            Console.WriteLine(oMsg.SenderName);
            Console.WriteLine(oMsg.ReceivedTime);
            Console.WriteLine(oMsg.Body);
            MessageBox.Show(a);*/

            try //U mnie dziala XD sprawdzcie czy u was tez
            {
                //getAllEmails(oItems);

                Outlook.Application myApp = new Outlook.Application();
                Outlook.NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
                Outlook.MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);


                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.Workbooks.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;


                workSheet.Cells[1, 1] = "Category";
                workSheet.Cells[1, 2] = "TIME: " + DateTime.Now.ToLongTimeString();
                workSheet.Cells[1, 3] = "Subject";
                workSheet.Cells[1, 4] = "Date-Time";
                workSheet.Cells[1, 5] = "SenderName";

                var row = 1;
                Outlook.MailItem newEmail = null;
                foreach (object collectionItem in oItems)
                {
                    
                    newEmail = collectionItem as Outlook.MailItem;
                    if (newEmail != null)
                    {
                        //row++;
                        if(newEmail.Categories == "Orange Category")
                        {
                            row++;
                            workSheet.Cells[row, 1] = newEmail.Categories;
                            workSheet.Cells[row, 3] = newEmail.Subject;
                            workSheet.Cells[row, 4] = newEmail.ReceivedTime;
                            workSheet.Cells[row, 5] = newEmail.SenderName;
                        }
                        if (newEmail.Categories == "Green Category")
                        {
                            row++;
                            workSheet.Cells[row, 1] = newEmail.Categories;
                            workSheet.Cells[row, 3] = newEmail.Subject;
                            workSheet.Cells[row, 4] = newEmail.ReceivedTime;
                            workSheet.Cells[row, 5] = newEmail.SenderName;
                        }
                        else { }
                    }
                    /*workSheet.Columns[1].AutoFit();//only for AutoFit
                    workSheet.Columns[2].AutoFit();
                    workSheet.Columns[3].AutoFit();
                    workSheet.Columns[4].AutoFit();
                    workSheet.Columns[5].AutoFit();*/
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
       
        public Ribbon1()
        {

         
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private static void getAllEmails(Outlook.Items oItems)
        {
            String c = "";
            Outlook.MailItem newEmail = null;
            foreach (object collectionItem in oItems)
            {
                newEmail = collectionItem as Outlook.MailItem;
                if (newEmail != null)
                {
                    c += "\n" + newEmail.ReceivedTime + "  " + newEmail.SenderName;
                }
            }
            MessageBox.Show(c);
        }

        #endregion
    }
}
