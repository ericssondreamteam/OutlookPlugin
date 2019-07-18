using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{

    public class DataObject
    {
        public List<string> inflow = new List<string>();
        public List<string> outflow = new List<string>();
        public List<string> inhands = new List<string>();
        public int inflowAmount = 0;
        public int outflowAmount = 0;
        public int inhandsAmount = 0;

        public void addNewItem(string n, List<bool> categoryList)
        {
            if (categoryList[0] == true)
            {
                inflowAmount++;
                inflow.Add(n);
            }
            if (categoryList[2] == true)
            {
                outflowAmount++;
                outflow.Add(n);
            }
            if (categoryList[1] == true)
            {
                inhandsAmount++;
                inhands.Add(n);
            }
        }
        public void ClearData()
        {
            inflow.Clear();
            outflow.Clear();
            inhands.Clear();
            inflowAmount = 0;
            outflowAmount = 0;
            inhandsAmount = 0;
        }
    }
}
