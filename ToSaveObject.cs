using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    class ToSaveObject
    {
        public List<string> inflow = new List<string>();
        public List<string> outflow = new List<string>();
        public List<string> inhands = new List<string>();
        public int inflowAmount=0;
        public int outflowAmount = 0;
        public int inhandsAmount = 0;

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
