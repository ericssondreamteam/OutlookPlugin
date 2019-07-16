using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    public class Mail
    {
        public String subject;
        public int conversationAmount;
        public DateTime recivedTime;
        public String category;
        public String ConversationID;
        public Mail(String subject, int conversationAmount, DateTime recivedTime, String category, String ConversationID)
        {
            this.subject = subject;
            this.conversationAmount = conversationAmount;
            this.recivedTime = recivedTime;
            this.category = category;
        }

    }
}
