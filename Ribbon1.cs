using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;


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
                if (!String.IsNullOrEmpty(email))
                {
                    var olapp = new Microsoft.Office.Interop.Outlook.Application();
                    Object selObject = olapp.ActiveExplorer().Selection[1];

                    if (selObject is Microsoft.Office.Interop.Outlook.MailItem)
                    {                    
                        Outlook.MailItem mailItem = this.GetSelectedItem();
                        if(mailItem != null)
                        {                                                                                                                  
                            mailItem.Recipients.Add(this.email);
                            mailItem.Send();                        
                        }   
                    }
                }
                else
                {
                    MessageBox.Show("Email isn't configured");
                }
            }
            catch (System.Exception ex)
            {                

            }
        }
        

        /// <summary>
        /// Forward the selected email as an attachment
        /// </summary>
        private void SendMailAsAttachment()
        {
            if (!String.IsNullOrEmpty(email))
            {
                var olapp = new Microsoft.Office.Interop.Outlook.Application(); 

                Outlook.MailItem mail = olapp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            
                mail.Subject = "ForwardIt";                                    
                //Add the configured email.
                mail.Recipients.Add(email);                
                mail.Attachments.Add(GetSelectedItem(), Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);                       
                mail.Send();
            }
            else
            {
                MessageBox.Show("Email isn't configured");
            }            
        }


        /// <summary>
        /// Get and return the selected item. If it isn't a mail item return null
        /// </summary>
        /// <returns></returns>
        private Outlook.MailItem GetSelectedItem()
        {
            var olapp = new Microsoft.Office.Interop.Outlook.Application();
            Object selObject = olapp.ActiveExplorer().Selection[1];

            if (selObject is Microsoft.Office.Interop.Outlook.MailItem)
            {
                return selObject as Outlook.MailItem;
            }
            else
            {
                return null;
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
            //MessageBox.Show("Forwarding Email");
            //ProcessMail();
            //MessageBox.Show("Attaching Email as Attachment");
            SendMailAsAttachment();

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
