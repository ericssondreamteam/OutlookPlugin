using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    public class Mail
    {
        public Mail()
        {

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
                return table.GetRowCount();
            }
            catch (Exception e)
            {
                OurDebug.AppendInfo("Blad w liczbie konwersacji; funkcja getConversationAmount()");
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
    }
}
