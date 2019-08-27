using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    class Settings
    {
        public Settings()
        {
            checkList = new List<bool>();
            checkList.Add(false);
            checkList.Add(false);
            checkList.Add(false);
        }
        static public string boxMailName { get; set; }
        static public string raportDate { get; set; }
        static public List<bool> checkList { get; set; }

        static public DialogResult ifWeDoRaport;
        static public string OutputRaportFileName;
    }
}
