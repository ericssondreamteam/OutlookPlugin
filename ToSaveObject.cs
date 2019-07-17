using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    class ToSaveObject
    {
        private List<string> inflow = new List<string>();
        private List<string> outflow = new List<string>();
        private List<string> inhands = new List<string>();
        private int inflowAmount=0;
        private int outflowAmount = 0;
        private int inhandsAmount = 0;

        public void addNewItem(string n,string k)
        {
            if (k == "inflow")
            {
                inflowAmount++;
                inflow.Add(n);
            }
            if (k == "outflow")
            {
                outflowAmount++;
                outflow.Add(n);
            }
            if (k == "inhands")
            {
                inhandsAmount++;
                inhands.Add(n);
            }                
        }
        private StringBuilder WriteInCorrextFomrat()
        {
            StringBuilder koncowyString = new StringBuilder();
            int i;
            koncowyString.Append("Inflow: " + inflowAmount + "\n");
            for (i = 0; i < inflow.Count; i++)
                koncowyString.Append("\t" + inflow[i] + "\n");
            koncowyString.Append("In-hands: " + inhandsAmount + "\n");
            for (i = 0; i < inhands.Count; i++)
                koncowyString.Append("\t" + inhands[i] + "\n");
            koncowyString.Append("Outflow: " +outflowAmount + "\n");
            for (i = 0; i < outflow.Count; i++)
                koncowyString.Append("\t" + outflow[i] + "\n");

            return koncowyString;

        }
        public void WriteToTxtFile( string path)
        {
            File.WriteAllText(path, WriteInCorrextFomrat().ToString());
        }

        public StringBuilder WriteInCorrectFormat(ToSaveObject subjects) 
        {
            StringBuilder endingString = new StringBuilder();
            endingString.Append("Inflow: " + subjects.inflowAmount + "\n");
            int i;
            for (i = 0; i < subjects.inflow.Count; i++)
                endingString.Append("\t" + subjects.inflow[i] + "\n");
            endingString.Append("In-hands: " + subjects.inhandsAmount + "\n");
            for (i = 0; i < subjects.inhands.Count; i++)
                endingString.Append("\t" + subjects.inhands[i] + "\n");
            endingString.Append("Outflow: " + subjects.outflowAmount + "\n");
            for (i = 0; i < subjects.outflow.Count; i++)
                endingString.Append("\t" + subjects.outflow[i] + "\n");

            return endingString;

        }

        public void WriteToTxtFile(StringBuilder toBeSaved, string path)
        {
            MessageBox.Show(path);
            File.WriteAllText(path, toBeSaved.ToString());
        }
    }
}
