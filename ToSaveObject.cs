using System.Collections.Generic;

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
    }
}
