using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Globalization;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using System.Diagnostics;
namespace OutlookAddIn1
{
    class EmailFunctions
    {
        public string adminMail = "NC Mailbox";
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

        public void choiceOfFileFormat(List<bool> checkList)
        {
            if (checkList[0])
            {
                OurDebug.Enable();
            }
            if (checkList[1])
            {
                Ribbon1.checkExcel = true;
            }
            if (checkList[2])
            {
                Ribbon1.checkWord = true;
            }
        }
        void EnumerateConversation(object item, Outlook.Conversation conversation, int i, List<bool> categoryList)
        {
            Outlook.SimpleItems items =
            conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    if (myItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                        myItem as Outlook.MailItem;
                        Outlook.Folder inFolder =
                        mailItem.Parent as Outlook.Folder;
                        string msg = mailItem.Subject + " in folder " + inFolder.Name + " Sender: " + mailItem.SenderName + " Date: " + mailItem.ReceivedTime;
                        if(i == 0)
                        {
                            if (mailItem.ReceivedTime > getInflowDate())
                            {
                                msg += " TYP: INFLOW";
                                categoryList[0] = true;
                            }
                        }
                        else
                        {
                            if (mailItem.SenderName.Equals(adminMail) && mailItem.ReceivedTime > getInflowDate())
                            {
                                msg += " TYP: IN HANDS";
                                categoryList[1] = true;
                            }
                        }
                            
                        Debug.WriteLine(msg);
                        OurDebug.AppendInfo(msg);
                        i++;
                    }
             
                    EnumerateConversation(myItem, conversation, i, categoryList);
                }
            }
        }
        public List<bool> selectCorrectEmailType(MailItem newEmail)
        {
            try
            {
                List<bool> categoryList = new List<bool>();
                categoryList.Add(false);
                categoryList.Add(false);
                categoryList.Add(false);
                int i = 0;

                Outlook.Conversation conv = newEmail.GetConversation();
                Debug.WriteLine("Conversation Items from Root:");
                Outlook.SimpleItems simpleItems = conv.GetRootItems();

                foreach (object item in simpleItems)
                {
                    try
                    {
                        if(item is Outlook.MailItem)
                        {
                            Outlook.MailItem mail = item as Outlook.MailItem;
                            Outlook.Folder inFolder = mail.Parent as Outlook.Folder;
                            string msg = mail.Subject + " in folder " + inFolder.Name + " Sender: " + mail.SenderName + " Date: " + mail.ReceivedTime;
                            if (mail.ReceivedTime > getInflowDate())
                            {
                                msg += " TYP: INFLOW";
                                categoryList[0] = true;
                            }
                            if (mail.SenderName.Equals(adminMail) && mail.ReceivedTime > getInflowDate())
                            {
                                msg += " TYP: IN HANDS";
                                categoryList[1] = true;
                            }
                            Debug.WriteLine(msg);
                            OurDebug.AppendInfo(msg);
                            i++;
                        }
                        EnumerateConversation(item, conv, i, categoryList);
                    }
                    catch(Exception e)
                    {
                        Debug.WriteLine("Exception in read children and set categories");
                        OurDebug.AppendInfo("Exception in read children and set categories");
                    }
                }
                if(categoryList[0] || categoryList[1])
                {
                    Debug.WriteLine("INFLOW: "+categoryList[0] + " INHANDS: " + categoryList[1] + " OUTFLOW: " + categoryList[2]);
                    Debug.WriteLine("----------------------------------------------");
                    OurDebug.AppendInfo("INFLOW: " + categoryList[0] + " INHANDS: " + categoryList[1] + " OUTFLOW: " + categoryList[2]);
                    OurDebug.AppendInfo("----------------------------------------------");
                    return categoryList;
                }
                categoryList[2] = true;
                Debug.WriteLine("INFLOW: " + categoryList[0] + " INHANDS: " + categoryList[1] + " OUTFLOW: " + categoryList[2]);
                Debug.WriteLine("----------------------------------------------");
                OurDebug.AppendInfo("INFLOW: " + categoryList[0] + " INHANDS: " + categoryList[1] + " OUTFLOW: " + categoryList[2]);
                OurDebug.AppendInfo("----------------------------------------------");
               
                return categoryList;
            }
            catch (Exception e)
            {
                OurDebug.AppendInfo("Blad w liczbie konwersacji; funkcja getConversationAmount()");
                Debug.WriteLine("Blad w liczbie konwersacji; funkcja getConversationAmount()");
                return null;
            }

        }
        public List<MailItem> emailsWithoutDuplicates(List<MailItem> emails)
        {
            try
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
            catch(Exception ex)
            {
                OurDebug.AppendInfo("Blad w usuwaniu duplikatow; brak dostepu do ConversationID");
                Debug.WriteLine("Blad w usuwaniu duplikatow; brak dostepu do ConversationID");
                return emails;
            }

        }
        public bool isMultipleCategoriesAndAnyOfTheireInterestedUs(string categories)
        {
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
                {   
                    if (!cat.Equals("noresponsenecessary") && !cat.Equals("unknown") && !cat.Equals("") && !cat.Equals("wow"))
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
        public int getOnlyEmailsForTwoWeeksAgo(int DebugForEachCounter, MailItem email1, Items oItems, int DebugCorrectEmailsCounter, List<MailItem> emails)
        {
            foreach (object collectionItem in oItems)
            {
                try
                {
                    DebugForEachCounter++;
                    email1 = collectionItem as MailItem;
                    if (email1 != null)
                    {
                        
                        OurDebug.AppendInfo("Email  ", DebugCorrectEmailsCounter.ToString(), ": ", email1.Subject, email1.ReceivedTime.ToString());

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
            return DebugForEachCounter;
        }
    }
}
