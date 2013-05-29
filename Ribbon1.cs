using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;


namespace ForwardIt
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private string email = string.Empty;


        public Ribbon1(string email)
        {
            this.email = email;
        }


        private void ProcessMail()
        {
            try
            {
                var olapp = new Microsoft.Office.Interop.Outlook.Application();
                Object selObject = olapp.ActiveExplorer().Selection[1];

                if (selObject is Microsoft.Office.Interop.Outlook.MailItem)
                {
                    Microsoft.Office.Interop.Outlook.MailItem mailItem = (selObject as Microsoft.Office.Interop.Outlook.MailItem);
                    var fwdmail = mailItem.Forward();
                    fwdmail.Recipients.Add(this.email);
                    fwdmail.Send();
                }
            }
            catch (System.Exception ex)
            {
                //todo log this.

            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ForwardIt.RibbonUIDef.xml");
        }

        #endregion

        #region Ribbon Callbacks        

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnClick(Office.IRibbonControl control)
        {
            //Process the mail.
            ProcessMail();
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

        #endregion
    }
}
