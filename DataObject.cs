using System;
using System.Collections.Generic;
using System.Diagnostics;

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
        public static int distance = 0;
        public static double percentage = 0.0;
        Debuger OurDebug;

        public DataObject(Debuger OurDebug)
        {
            this.OurDebug = OurDebug;
        }

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
        public void lastTuning()
        {
            removeReAndFW(inflow);
            removeReAndFW(outflow);
            removeReAndFW(inhands);

            checkStringSimilarity(inflow);
            checkStringSimilarity(outflow);
            checkStringSimilarity(inhands);
        }
        private void removeReAndFW(List<string> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = i + 1; j < list.Count; j++)
                {
                    if (list[i].Trim().ToLower().StartsWith("re:") || list[i].Trim().ToLower().StartsWith("fw:"))
                    {
                        list[i] = list[i].Substring(4);
                    }
                }
            }
        }

        public double similarity(string s1, string s2)
        {

            string longer = s1, shorter = s2;
            if (s1.Length < s2.Length)
            { // longer should always have greater length
                longer = s2; shorter = s1;
            }
            int longerLength = longer.Length;
            if (longerLength == 0) { return 1.0; /* both strings are zero length */ }
            distance = getLevenshteinDistance(longer, shorter);
            return (longerLength - distance) / (double)longerLength;

        }

        public int getLevenshteinDistance(string s1, string s2)
        {
            s1 = s1.ToLower();
            s2 = s2.ToLower();

            int[] costs = new int[s2.Length + 1];
            for (int i = 0; i <= s1.Length; i++)
            {
                int lastValue = i;
                for (int j = 0; j <= s2.Length; j++)
                {
                    if (i == 0)
                        costs[j] = j;
                    else
                    {
                        if (j > 0)
                        {
                            int newValue = costs[j - 1];
                            if (s1[i - 1] != s2[j - 1])
                                newValue = Math.Min(Math.Min(newValue, lastValue),
                                        costs[j]) + 1;
                            costs[j - 1] = lastValue;
                            lastValue = newValue;
                        }
                    }
                }
                if (i > 0)
                    costs[s2.Length] = lastValue;
            }
            Console.WriteLine("DEBUG----------> " + costs[s2.Length]);
            return costs[s2.Length];
        }

        public List<String> checkStringSimilarity(List<String> emails)
        {
            try
            {
                for (int i = 0; i < emails.Count; i++)
                {
                    for (int j = i + 1; j < emails.Count; j++)
                    {
                        percentage = similarity(emails[i], emails[j]);
                        Debug.WriteLine("------> checkStringSimilarity() <------ PERCENTAGE: "
                            + percentage + " DISTANCE: " + distance);
                        //TERAZ CAŁA LOGIKA if else if oraz else

                        if (emails[i].StartsWith("https") || emails[j].StartsWith("https"))
                        {
                            if (percentage == 1.0)
                            {
                                emails.RemoveAt(j);
                                j--;

                            }
                            else
                            {
                                //NIC
                            }
                        }
                        else if ((emails[i].Length < 35 || emails[j].Length < 35) && distance <= 4 && percentage >= 0.87)
                        {
                            emails.RemoveAt(j);
                            j--;
                        }
                        else if ((emails[i].Length > 35 || emails[j].Length > 35) && distance <= 3 && percentage >= 0.92)
                        {
                            emails.RemoveAt(j);
                            j--;
                        }
                        else if ((emails[i].Length > 35 || emails[j].Length > 35) && distance <= 4 && percentage >= 0.95)
                        {
                            emails.RemoveAt(j);
                            j--;
                        }
                        else
                        {
                            //NIC
                        }

                    }
                }
                return emails;
            }
            catch (Exception ex)
            {
                OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "DataObject.cs line:176. Blad w usuwaniu duplikatow; Sprawdzanie poprawnosci procentowej.\n", ex.Message, "\n", ex.StackTrace);
                Debug.WriteLine("Blad w usuwaniu duplikatow; Sprawdzanie poprawnosci procentowej ");
                return emails;
            }

        }
    }
}
