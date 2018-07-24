using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
//using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Tricks
{
    public partial class ThisAddIn
    {
        private static readonly Framework.Ribbon Ribbon = new Framework.Ribbon("Tricks.Views.Ribbon.xml");

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Ribbon.ClickButton += Ribbon_ClickButton;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Ribbon.ClickButton -= Ribbon_ClickButton;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Ribbon;
        }

        void Ribbon_ClickButton(object sender, Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "BtnCreateSheetNameList":
                    {
                        var buf = new Excel.Helpers.WorksheetRAM(Globals.ThisAddIn.Application);
                        buf.WriteOnNewSheet(buf.CreateSheetNameTbl());
                        break;
                    }
                default:
                    break;
            }
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
