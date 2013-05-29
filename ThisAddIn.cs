
using Microsoft.Win32;
using System;
namespace ForwardIt
{
    public partial class ThisAddIn
    {


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1(GetEmailFromReg());
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.PluginInit();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {


        }

        private void PluginInit()
        {
            //init code here
        }



        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private string GetEmailFromReg()
        {                                    
            RegistryKey sk1 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\OutlookSample");            
            string keyname = "email";
            
            if ( sk1 == null )
            {
                return null;
            }
            else
            {
                try 
                {                    
                    return (string)sk1.GetValue(keyname);
                }
                catch (Exception ex)
                {                                        
                    return null;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
