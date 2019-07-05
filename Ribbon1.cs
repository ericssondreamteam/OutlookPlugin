
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

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

            /*foreach(Outlook.MailItem mail in oItems)
            {
                MessageBox.Show(mail.ReceivedTime + mail.SenderName);
            }*/

            /*var lista = new System.Collections.Generic.List<Outlook.MailItem>();
            lista.Add(oMsg);*/

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

            try
            {
                Outlook.MailItem oMsg = (Outlook.MailItem)oItems.GetFirst();
                String a = oMsg.ReceivedTime + " " + oMsg.SenderName;
                MessageBox.Show(a);

                oMsg = (Outlook.MailItem)oItems.GetNext();
                a = oMsg.ReceivedTime + " " + oMsg.SenderName;
                MessageBox.Show(a);

                oMsg = (Outlook.MailItem)oItems.GetLast();
                a = oMsg.ReceivedTime + " " + oMsg.SenderName;
                MessageBox.Show(a);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
            


            /*String expMessage = "cos XD ";
            String itemMessage = "Item is unknown.";
            try
            {
                if(oApp.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObj = oApp.ActiveExplorer().Selection[1];
                    if(selObj is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = (selObj as Outlook.MailItem);
                        itemMessage = "The item is an e-mail message." +
                        " The subject is " + mailItem.Subject + ".";
                        mailItem.Display(false);
                    }
                    else if (selObj is Outlook.ContactItem)
                    {
                        Outlook.ContactItem contactItem =
                            (selObj as Outlook.ContactItem);
                        itemMessage = "The item is a contact." +
                            " The full name is " + contactItem.Subject + ".";
                        contactItem.Display(false);
                    }
                    else if (selObj is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem apptItem =
                            (selObj as Outlook.AppointmentItem);
                        itemMessage = "The item is an appointment." +
                            " The subject is " + apptItem.Subject + ".";
                    }
                    else if (selObj is Outlook.TaskItem)
                    {
                        Outlook.TaskItem taskItem =
                            (selObj as Outlook.TaskItem);
                        itemMessage = "The item is a task. The body is "
                            + taskItem.Body + ".";
                    }
                    else if (selObj is Outlook.MeetingItem)
                    {
                        Outlook.MeetingItem meetingItem =
                            (selObj as Outlook.MeetingItem);
                        itemMessage = "The item is a meeting item. " +
                             "The subject is " + meetingItem.Subject + ".";
                    }
                }
                expMessage = expMessage + itemMessage;
            }
            catch (Exception e)
            {
                expMessage = e.Message;
            }
            MessageBox.Show(expMessage);*/
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

        #endregion
    }
}
