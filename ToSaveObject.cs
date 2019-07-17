using System.Collections.Generic;

namespace OutlookAddIn1
{
    class ToSaveObject
    {
        public List<string> inflow = new List<string>();
        public List<string> outflow = new List<string>();
        public List<string> inhands = new List<string>();

        public void addNewItem(string n,string k)
        {
            if (k == "inflow")
                inflow.Add(n);
            if (k == "outflow")
                outflow.Add(n);
            if (k == "inhands")
                inhands.Add(n);
        }
    }
}
