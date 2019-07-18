using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;

namespace OutlookAddIn1
{
    class ToSaveObject
    {
        public List<string> inflow = new List<string>();
        public List<string> outflow = new List<string>();
        public List<string> inhands = new List<string>();
        public int inflowAmount = 0;
        public int outflowAmount = 0;
        public int inhandsAmount = 0;
        
        private StringBuilder WriteInCorrextFomrat()
        {
            StringBuilder endingString = new StringBuilder();
            int i;
            endingString.Append("Inflow: " + inflowAmount + "\n");
            for (i = 0; i < inflow.Count; i++)
                endingString.Append("\t" + inflow[i] + "\n");
            endingString.Append("In-hands: " + inhandsAmount + "\n");
            for (i = 0; i < inhands.Count; i++)
                endingString.Append("\t" + inhands[i] + "\n");
            endingString.Append("Outflow: " + outflowAmount + "\n");
            for (i = 0; i < outflow.Count; i++)
                endingString.Append("\t" + outflow[i] + "\n");

            return endingString;

        }
        public void WriteToTxtFile(string path)
        {
            File.WriteAllText(path, WriteInCorrextFomrat().ToString());
            inhands.Clear();
            inflow.Clear();
            outflow.Clear();
            inflowAmount = 0;
            outflowAmount = 0;
            inhandsAmount = 0;
        }
        
    }
}
