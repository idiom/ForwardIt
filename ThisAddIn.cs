
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


        private void CleanUp()
        {
            //cleanup code here
        }



        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private string GetEmailFromReg()
        {                                                
            RegistryKey root   = Registry.LocalMachine;
            RegistryKey subkey = root.OpenSubKey(@"SOFTWARE\OutlookSample");            
            string keyname = "Email";
            return (subkey == null ? null : subkey.GetValue(keyname) as string);
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
