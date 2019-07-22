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
       
        public EmailFunctions(Debuger OurDebug,string mailName)
        {
            this.OurDebug = OurDebug;
            adminMail = mailName;
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
            SimpleItems items = conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    if (myItem is Outlook.MailItem)
                    {
                        MailItem mailItem = myItem as MailItem;
                        Folder inFolder = mailItem.Parent as Folder;
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

                Conversation conv = newEmail.GetConversation();
                Debug.WriteLine("Conversation Items from Root:");
                SimpleItems simpleItems = conv.GetRootItems();

                foreach (object item in simpleItems)
                {
                    try
                    {
                        if(item is MailItem)
                        {
                            MailItem mail = item as MailItem;
                            Folder inFolder = mail.Parent as Folder;
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
                        Debug.WriteLine("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Exception in read children and set categories");
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
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Blad w liczbie konwersacji; funkcja getConversationAmount()\n",e.Message,"\n",e.StackTrace);
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
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Blad w usuwaniu duplikatow; brak dostepu do ConversationID\n",ex.Message,"\n",ex.StackTrace);
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


            for (int i = 0; i < emails.Count - 1; i++)
            {
                mailSubject1 = emails[i].Subject.Trim().Replace(" ","").ToLower();
                for (int j = i + 1; j < emails.Count; j++)
                {
                    mailSubject2 = emails[j].Subject.Trim().Replace(" ", "").ToLower();                    
                    //string mailSubject1;
                    //string mailSubject2;
                    if(mailSubject2.Length > mailSubject1.Length)
                    {
                        var a = mailSubject1.Substring(mailSubject1.Length / 2);
                        if (mailSubject2.Contains(a))
                        {
                            emails.RemoveAt(j);
                            j--;
                        }
                    }
                    else
                    {
                        var a = mailSubject2.Substring(mailSubject2.Length / 2);
                        if (mailSubject1.Contains(a))
                        {
                            emails.RemoveAt(j);
                            j--;
                        }
                    }
                   

                }

            }

            return emails;

        }

        private static int levenshtein(String s, String t)
        {
            int i, j, m, n, cost;
            int[,] d;

            m = s.Length;
            n = t.Length;

            d = new int[m + 1, n + 1];

            for (i = 0; i <= m; i++)
                d[i, 0] = i;
            for (j = 1; j <= n; j++)
                d[0, j] = j;

            for (i = 1; i <= m; i++)
            {
                for (j = 1; j <= n; j++)
                {
                    if (s[i - 1] == t[j - 1])
                        cost = 0;
                    else
                        cost = 1;

                    d[i, j] = Math.Min(d[i - 1, j] + 1,   /* remove */
                    Math.Min(d[i, j - 1] + 1,         /* insert */
                    d[i - 1, j - 1] + cost));        /* change */
                }
            }

            return d[m, n];
        }

        public static double obliczPodobienstwo(String lancuchPierwszy, String lancuchDrugi)
        {
            // obliczamy i zwracamy podobieństwo łańcuchów
            return (1.0 / (1.0 + levenshtein(lancuchPierwszy, lancuchDrugi)));
        }

        public int getOnlyEmailsForTwoWeeksAgo(int DebugForEachCounter, MailItem email1, Items oItems, int DebugCorrectEmailsCounter, List<MailItem> emails)
        {
            foreach (object collectionItem in oItems)
            {
                try
                {
                    
                    email1 = collectionItem as MailItem;
                    if (email1 != null)
                    {
                        DebugForEachCounter++;
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
