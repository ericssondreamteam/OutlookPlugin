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
       public int getConversationAmount(MailItem newEmail)
        {
            try
            {
                Outlook.Conversation conv = newEmail.GetConversation();
                Outlook.Table table = conv.GetTable();
                Debug.WriteLine("Pobieramy maile z conwersacji NOWA");
                Debug.WriteLine("+++++++++++++++++++++++++++++++++++++");
                Array tableArray = table.GetArray(table.GetRowCount()) as Array;
                for (int i = 0; i <= tableArray.GetUpperBound(0); i++)
                {
                    for (int j = 0; j <= tableArray.GetUpperBound(1); j++)
                    {
                        Debug.WriteLine(tableArray.GetValue(i, j));
                    }
                }

                Debug.WriteLine("+++++++++++++++++++++++++++++++++++++");
                return table.GetRowCount();
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
                if (getConversationAmount(newEmail) > 1 && newEmail.ReceivedTime > getInflowDate()) //in hands
                {
                    typ = 1;
                }
                else if (newEmail.ReceivedTime > getInflowDate()) //inflow
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
