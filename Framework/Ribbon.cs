using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace Tricks.Framework
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private string resourceName;

        public event Action<object, Office.IRibbonControl> ClickButton;
        public event Action<object, Exception> ExceptionOccured;

        public Ribbon(string resourceName)
        {
            this.resourceName = resourceName;
        }

        public void FireClickButton(Office.IRibbonControl control)
        {
            if (null == ClickButton)
            {
                return;
            }
            ClickButton.Invoke(this, control);
        }

        public void FireExceptionOccured(Exception ex)
        {
            if (null == ExceptionOccured)
            {
                return;
            }
            ExceptionOccured.Invoke(this, ex);
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText();
        }

        #endregion

        #region リボンのコールバック
        //ここにコールバック メソッドを作成します。コールバック メソッドの追加の詳細については、ソリューション エクスプローラーでリボンの XML アイテムを選択し、F1 キーを押してください

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void Button_OnAction(Office.IRibbonControl control)
        {
            if (null == ClickButton)
            {
                return;
            }
            try
            {
                FireClickButton(control);
            }
            catch (Exception ex)
            {
                FireExceptionOccured(ex);
            }
        }

        #endregion

        #region ヘルパー

        private string GetResourceText()
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
