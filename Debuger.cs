using System;
using System.Text;


namespace OutlookAddIn1
{
    class Debuger
    {
        private StringBuilder DebugerMsg = new StringBuilder();
        private bool debugerStatus = false;
        public bool IsEnable()
        {
            return debugerStatus;
        }
        public void Enable()
        {
            debugerStatus = true;
        }
        public void Disable()
        {
            debugerStatus = false;
        }
        public void AppendInfo(params string[] info)
        {
            if (debugerStatus)
            {
                foreach (string s in info)
                {
                    DebugerMsg.Append(s);
                    DebugerMsg.Append(" ");
                }
                DebugerMsg.Append("\n");
            }
        }
        public void SaveDebugInfoToFile(string path)
        {
            if (debugerStatus)
            {
                try
                {
                    System.IO.File.WriteAllText(path, DebugerMsg.ToString());
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError during save debug info\n");
                    Console.WriteLine(e.Message + "\n");
                    Console.WriteLine(e.StackTrace);
                }
            }
        }
    }
}
