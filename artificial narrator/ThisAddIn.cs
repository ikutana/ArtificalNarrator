using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace artificial_narrator
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowActivate += ChangeButtonState;
            this.Application.WindowDeactivate += ChangeButtonState;
        }

        void ChangeButtonState(PowerPoint.Presentation Pres = null, PowerPoint.DocumentWindow Wn = null)
        {
            var TheActivePresentation = Pres;
            Globals.Ribbons.Ribbon.InsertNarration.Enabled = (TheActivePresentation!= null);  
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
