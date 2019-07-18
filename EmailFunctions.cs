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
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using System.Diagnostics;
using System.Collections;
using System.Threading;
using System.Text;

namespace OutlookAddIn1
{
    class EmailFunctions
    {
        Debuger OurDebug;
        public EmailFunctions(Debuger OurDebug)
        {
            this.OurDebug = OurDebug;
        }
        public static DateTime GetFirstDayOfWeek(DateTime dayInWeek)
        {
            CultureInfo defaultCultureInfo = CultureInfo.CurrentCulture;
            return GetFirstDayOfWeek(dayInWeek, defaultCultureInfo);
        }

        public static DateTime GetFirstDayOfWeek(DateTime dayInWeek, CultureInfo cultureInfo)
        {
            DayOfWeek firstDay = cultureInfo.DateTimeFormat.FirstDayOfWeek;
            DateTime firstDayInWeek = dayInWeek.Date;
            while (firstDayInWeek.DayOfWeek != firstDay)
                firstDayInWeek = firstDayInWeek.AddDays(-1);
            return firstDayInWeek;
        }
        public DateTime getInflowDate()
        {
            DateTime today = GetFirstDayOfWeek(DateTime.Today);
            today = today.AddDays(-2).AddHours(17);
            return today;
        }

        /*void DemoConversation()
        {
            object selectedItem =
            Application.ActiveExplorer().Selection[1];
            // This example uses only 
            // MailItem. Other item types such as 
            // MeetingItem and PostItem can participate 
            // in the conversation. 
            if (selectedItem is Outlook.MailItem)
            {
                // Cast selectedItem to MailItem. 
                Outlook.MailItem mailItem =
                selectedItem as Outlook.MailItem;
                // Determine the store of the mail item. 
                Outlook.Folder folder = mailItem.Parent
                as Outlook.Folder;
                Outlook.Store store = folder.Store;
                if (store.IsConversationEnabled == true)
                {
                    // Obtain a Conversation object. 
                    Outlook.Conversation conv =
                    mailItem.GetConversation();
                    // Check for null Conversation. 
                    if (conv != null)
                    {
                        // Obtain Table that contains rows 
                        // for each item in the conversation. 
                        Outlook.Table table = conv.GetTable();
                        Debug.WriteLine("Conversation Items Count: " +
                        table.GetRowCount().ToString());
                        Debug.WriteLine("Conversation Items from Table:");
                        while (!table.EndOfTable)
                        {
                            Outlook.Row nextRow = table.GetNextRow();
                            Debug.WriteLine(nextRow["Subject"]
                            + " Modified: "
                            + nextRow["LastModificationTime"]);
                        }
                        Debug.WriteLine("Conversation Items from Root:");
                        // Obtain root items and enumerate the conversation. 
                        Outlook.SimpleItems simpleItems
                        = conv.GetRootItems();
                        foreach (object item in simpleItems)
                        {
                            // In this example, only enumerate MailItem type. 
                            // Other types such as PostItem or MeetingItem 
                            // can appear in the conversation. 
                            if (item is Outlook.MailItem)
                            {
                                Outlook.MailItem mail = item
                                as Outlook.MailItem;
                                Outlook.Folder inFolder =
                                mail.Parent as Outlook.Folder;
                                string msg = mail.Subject
                                + " in folder " + inFolder.Name;
                                Debug.WriteLine(msg);
                            }
                            // Call EnumerateConversation 
                            // to access child nodes of root items. 
                            EnumerateConversation(item, conv);
                        }
                    }
                }
            }
        }*/


        void EnumerateConversation(object item,
         Outlook.Conversation conversation)
        {
            Outlook.SimpleItems items =
            conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    // In this example, only enumerate MailItem type. 
                    // Other types such as PostItem or MeetingItem 
                    // can appear in the conversation. 
                    if (myItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                        myItem as Outlook.MailItem;
                        Outlook.Folder inFolder =
                        mailItem.Parent as Outlook.Folder;
                        string msg = mailItem.Subject
                        + " in folder " + inFolder.Name + " Sender: " + mailItem.SenderName
                        + " Date: " + mailItem.ReceivedTime;
                        if (mailItem.SenderName == "Karol Lasek" && mailItem.ReceivedTime > getInflowDate())
                        {
                            msg += " TYP: IN HANDS";
                        }
                        Debug.WriteLine(msg);
                    }
                    // Continue recursion. 
                    EnumerateConversation(myItem, conversation);
                }
            }
        }


        public int getConversationAmount(MailItem newEmail)
        {
            try
            {
                Outlook.Conversation conv = newEmail.GetConversation();
                Debug.WriteLine("Conversation Items from Root:");
                Outlook.SimpleItems simpleItems
                = conv.GetRootItems();
                foreach (object item in simpleItems)
                {
                    if (item is Outlook.MailItem)
                    {
                        Outlook.MailItem mail = item
                        as Outlook.MailItem;
                        Outlook.Folder inFolder =
                        mail.Parent as Outlook.Folder;
                        string msg = mail.Subject
                        + " in folder " + inFolder.Name + " Sender: " + mail.SenderName
                        + " Date: " + mail.ReceivedTime;
                        if (mail.ReceivedTime > getInflowDate())
                        {
                            msg += " TYP: INFLOW";
                        }
                        if(mail.SenderName == "Karol Lasek" && mail.ReceivedTime > getInflowDate())
                        {
                            msg += " TYP: IN HANDS";
                        }
                        
                        Debug.WriteLine(msg);
                    }
                    EnumerateConversation(item, conv);
                }
                return 1;
            }
            catch (Exception e)
            {
                OurDebug.AppendInfo("Blad w liczbie konwersacji; funkcja getConversationAmount()");
                Debug.WriteLine("Blad w liczbie konwersacji; funkcja getConversationAmount()");
                return 0;
            }

        }
        public int selectCorrectEmailType(MailItem newEmail)
        {
            int typ = 0;
            if (newEmail.Categories != null)
            {
                /*if (getConversationAmount(newEmail) > 1 && newEmail.ReceivedTime > getInflowDate()) //in hands
                {
                    typ = 1;
                }*/
                if (newEmail.ReceivedTime > getInflowDate()) //inflow
                {
                    typ = 2;
                }
                else if ((newEmail.ReceivedTime > getInflowDate().AddDays(-7)) && (newEmail.ReceivedTime < getInflowDate())) //outflow
                {
                    typ = 3;
                }
                if (typ == 1) //inflow + in hands
                {
                    typ = 4;
                }
            }
            OurDebug.AppendInfo("Nadany typ:", typ.ToString());
            return typ;
        }

        public List<MailItem> emailsWithoutDuplicates(List<MailItem> emails)
        {
            for (int i = 0; i < emails.Count; i++)
            {
                for (int j = i + 1; j < emails.Count; j++)
                {
                    if (emails[i].ConversationID.Equals(emails[j].ConversationID))
                    {

                        emails.RemoveAt(j);
                        j--;
                    }
                }
            }
            return emails;
        }
        public bool isMultipleCategoriesAndAnyOfTheireInterestedUs(string categories)
        {
            OurDebug.AppendInfo("Categories start:", categories);
            if (categories is null)
            {
                return false;
            }
            else
            {
                categories = categories.Trim();
                categories = categories.Replace(" ", "");
                categories = categories.ToLower();
                OurDebug.AppendInfo("Categories after trim and repalce and lower:", categories);
                string[] categoriesList = categories.Split(',');
                foreach (var cat in categoriesList)
                {   //No Response Necessary    or    Unknown     No Response Necessary, Unknown
                    if (!cat.Equals("noresponsenecessary") && !cat.Equals("unknown") && !cat.Equals(""))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        public List<MailItem> removeDuplicateOneMoreTime(List<MailItem> emails)
        {
            string mailSubject1;
            string mailSubject2;
            for (int i = 0; i < emails.Count - 1; i++)
            {
                mailSubject1 = emails[i].Subject;
                mailSubject1 = mailSubject1.Trim();
                mailSubject1 = mailSubject1.Replace(" ", "");
                mailSubject1 = mailSubject1.ToLower();
                if (mailSubject1.Substring(0, 3).Equals("re:") || mailSubject1.Substring(0, 3).Equals("fw:"))
                    mailSubject1 = mailSubject1.Substring(3);

                for (int j = i + 1; j < emails.Count; j++)
                {
                    mailSubject2 = emails[j].Subject;
                    mailSubject2 = mailSubject2.Trim();
                    mailSubject2 = mailSubject2.Replace(" ", "");
                    mailSubject2 = mailSubject2.ToLower();
                    if (mailSubject2.Substring(0, 3).Equals("re:") || mailSubject2.Substring(0, 3).Equals("fw:"))
                        mailSubject2 = mailSubject2.Substring(3);

                    if (mailSubject1.Equals(mailSubject2))

                    {
                        emails.RemoveAt(j);
                        j--;
                    }
                }
            }
            return emails;
        }
        public void getOnlyEmailsForTwoWeeksAgo(int DebugForEachCounter, MailItem email1, Items oItems, int DebugCorrectEmailsCounter, List<MailItem> emails)
        {
            foreach (object collectionItem in oItems)
            {
                try
                {
                    DebugForEachCounter++;
                    email1 = collectionItem as MailItem;
                    if (email1 != null)
                    {
                        //Save mails
                        OurDebug.AppendInfo("Email  ", DebugCorrectEmailsCounter.ToString(), ": ", email1.Subject, email1.ReceivedTime.ToString());

                        //Add to list of mails
                        if (email1.ReceivedTime > getInflowDate().AddDays(-7))
                        {
                            DebugCorrectEmailsCounter++;
                            emails.Add(email1);
                        }
                        else
                            break;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show("Some error occured during first analysis\nIf You turn on debugger please go there");
                    OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "FIRST TRY CATCH\n", "eMail number:", DebugCorrectEmailsCounter.ToString(), "\n", e.Message, "\n", e.StackTrace);
                }
            }
        }
    }
}
