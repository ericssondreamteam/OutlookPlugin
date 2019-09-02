using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {

        }

        public void OnTableButton(Office.IRibbonControl control)
        {
            Settings set = new Settings();            
            try
            {
                string OutputRaportFileName = "Raport_" + DateTime.Now.ToString("dd_MM_yyyy");
                Form1 form3 = new Form1(ref OutputRaportFileName);
                form3.ShowDialog();
                if (Settings.ifWeDoRaport == DialogResult.OK)
                {
                    Loading waitingScreen = new Loading();
                    waitingScreen.ShowDialog();
                }
                else
                {
                    //OPERATION CANCELED
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Some error occured during second analysis\nIf You turn on debugger please go there");
                Loading.OurDebug.AppendInfo("!!!!!!!!************ERROR***********!!!!!!!!!!\n", "Ribbon1.cs line:135. SECOND TRY CATCH\n", e.Message, "\n", e.StackTrace);
            }
            finally
            {
                if (Loading.OurDebug.IsEnable())
                {
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    path += "\\DebugInfoRaportPlugin.txt";
                    Loading.OurDebug.SaveDebugInfoToFile(path);
                    Loading.fullInfoBox += "\n\nYour debug file is saved: DebugInfoRaportPlugin.txt";
                    Loading.OurDebug.Disable();
                }
                Form2 summary = new Form2();
                if(Settings.ifWeDoRaport==DialogResult.OK)
                    summary.Show();

                Loading.OurData.ClearData();
                Loading.DebugForEachCounter = 0;
                Loading.checkExcel = false;
                Loading.checkWord = false;
                Loading.fullInfoBox = String.Empty;
            }
        }

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.Ribbon1.xml");
        }
        #endregion
        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        #endregion
        #region Helpers
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }



        private static void getAllEmails(Outlook.Items oItems)
        {
            String c = "";
            Outlook.MailItem newEmail = null;
            foreach (object collectionItem in oItems)
            {
                newEmail = collectionItem as Outlook.MailItem;
                if (newEmail != null)
                {
                    c += "\n" + newEmail.ReceivedTime + "  " + newEmail.SenderName;
                }
            }
            MessageBox.Show(c);
        }
        #endregion
    }
}
