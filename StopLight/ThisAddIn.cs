using System;
using Office = Microsoft.Office.Core;

namespace StopLight
{
    public partial class ThisAddIn
    {
        internal Ribbon Ribbon = null;
        internal HighlightManager HighlightManager = new HighlightManager();

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new Ribbon();
            return Ribbon;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            detectLanguage();
        }
        

        private void detectLanguage()
        {
            //language detection for Word document
            //also calling in Ribbon constructor
            //this works in Windows
            //not sure if this is grabbing the locale of the application or the Windows system
            //NEED TO TEST ON iOS BEFORE PUBLISHING
            long languageID = this.Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            //can find language ID
            //https://support.microsoft.com/en-us/kb/221435
            //https://msdn.microsoft.com/en-us/goglobal/bb964664.aspx?f=255&MSPPError=-2147217396
            long Japanese = 1041;
            //long English = 1033;

            if (languageID == Japanese)
            {
                Strings.Japanese();
            }
            else
            {
                //default is English
                //do not need to change text
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
