using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using UserControl = System.Windows.Forms.UserControl;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;
using System.Diagnostics;

namespace TreeViewer
{

    public class PaneManager {
        static Dictionary<string, CustomTaskPane> panes = new Dictionary<string, CustomTaskPane>();
        public static CustomTaskPane GetPane(string title, Func<UserControl> creator) {
            string key = Globals.ThisAddIn.Application.Hwnd.ToString();
            if (!panes.ContainsKey(key)) {
                panes[key] = Globals.ThisAddIn.CustomTaskPanes.Add(creator(), title);
            }
            return panes[key];
        }
    }

    public partial class ThisAddIn
    {

        //void Application_WorkbookOpen(Excel.Workbook wb) {
        //    myTreeViewer.UpdateTreeView();
        //}

        void Application_WorkbookActivate(Excel.Workbook wb) {
            Debug.Print("Activate {0}", wb.Name);
            var pane = PaneManager.GetPane("Switch", () => new TreeViewer());
            pane.Visible = true;
            var tree = pane.Control as TreeViewer;
            tree.UpdateTreeView();
        }

        void Application_WorksheetUpdate(Excel.Workbook wb, object obj) {
            var sheet = obj as Excel.Worksheet;
            Debug.Print("Activate {0}", sheet.Name);
            var pane = PaneManager.GetPane("Switch", () => new TreeViewer());
            pane.Visible = true;
            var tree = pane.Control as TreeViewer;
            tree.UpdateTreeView();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WorkbookActivate +=
                new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
            Application.WorkbookNewSheet +=
                new Excel.AppEvents_WorkbookNewSheetEventHandler(Application_WorksheetUpdate);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
