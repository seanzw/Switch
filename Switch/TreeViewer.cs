using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TreeViewer {

    public class APIUtils {
        public static IntPtr GetXLMainWindowHandle(IntPtr DesktopHandle) {
            return FindMainWindowInProcess(DesktopHandle, String.Empty, "XLMAIN");
        }

        internal static IntPtr FindMainWindowInProcess(IntPtr HWNDParent, string WindowName, string WindowClass) {
            IntPtr FoundWindow = IntPtr.Zero;

            string FindClass = String.Empty;
            string FindName = String.Empty;
            uint WindowProcessID;

            IntPtr tempWindow = GetWindow(HWNDParent, GW_CHILD);
            while ((int)tempWindow > 0) {
                FindName = GetWindowText(tempWindow);
                FindClass = GetClassName(tempWindow);

                bool CompareClass = ((FindClass.IndexOf(WindowClass) >= 0) || (WindowClass == String.Empty));
                bool CompareCaption = ((FindName.IndexOf(WindowName) >= 0) || (WindowName == String.Empty));

                if (CompareClass && CompareCaption) {
                    GetWindowThreadProcessId(tempWindow, out WindowProcessID);
                    if (GetCurrentProcessId() == WindowProcessID) {
                        FoundWindow = tempWindow;
                        if (IsWindowVisible(FoundWindow)) {
                            break;
                        }
                    }
                }
                tempWindow = GetWindow(tempWindow, GW_HWNDNEXT);
            }
            return FoundWindow;
        }

        [DllImport("user32.dll")]
        public static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetWindow(IntPtr hwnd, int uCmd);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hwnd,
            [In][Out]StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern int GetClassName(IntPtr hwnd,
            [In][Out]StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();

        public static string GetWindowText(IntPtr handle) {
            System.Text.StringBuilder lpBuffer = new System.Text.StringBuilder(255);
            if (GetWindowText(handle, lpBuffer, 255) > 0) {
                return lpBuffer.ToString();
            } else
                return "";
        }

        public static string GetClassName(IntPtr handle) {
            System.Text.StringBuilder className = new System.Text.StringBuilder(255);
            if (GetClassName(handle, className, 255) > 0) {
                return className.ToString();
            } else
                return "";
        }

        public const int GW_CHILD = 5;
        public const int GW_HWNDNEXT = 2;

    }

    public partial class TreeViewer : UserControl {

        public TreeViewer() {
            InitializeComponent();
            UpdateTreeView();
        }

        public void UpdateTreeView() {
            this.treeView1.BeginUpdate();
            this.treeView1.Nodes.Clear();
            //foreach (var obj in Globals.ThisAddIn.Application.Workbooks) {
            //    var wb = obj as Microsoft.Office.Interop.Excel.Workbook;
            //    var node = this.treeView1.Nodes.Add(wb.Name);
            //    foreach (var ss in wb.Worksheets) {
            //        var sheet = ss as Microsoft.Office.Interop.Excel.Worksheet;
            //        node.Nodes.Add(sheet.Name);
            //    }
            //    node.ExpandAll();
            //}
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var node = this.treeView1.Nodes.Add(wb.Name);
            foreach (var ss in wb.Worksheets) {
                var sheet = ss as Microsoft.Office.Interop.Excel.Worksheet;
                node.Nodes.Add(sheet.Name);
            }
            node.ExpandAll();
            this.treeView1.EndUpdate();
        }

        private void SetFocusToExcel() {
            var desk = APIUtils.GetDesktopWindow();
            APIUtils.SetForegroundWindow(desk);
            IntPtr h = APIUtils.GetXLMainWindowHandle(APIUtils.GetDesktopWindow());
            APIUtils.SetForegroundWindow(new IntPtr(Globals.ThisAddIn.Application.Hwnd));
        }


        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e) {

            // Set the focus back.
            //SetFocusToExcel();

            //MessageBox.Show(e.Node.Text);
            //string wbName = "";
            //if (e.Node.Level == 0) {
            //    // This is the workbook level node.
            //    wbName = e.Node.Text;
            //} else if (e.Node.Level == 1) {
            //    // This is the sheet level node.
            //    wbName = e.Node.Parent.Text;
            //}

            string wbName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;

            if (wbName != Globals.ThisAddIn.Application.ActiveWorkbook.Name) {
                var found = false;
                foreach (var o in Globals.ThisAddIn.Application.Workbooks) {
                    var wb = o as Microsoft.Office.Interop.Excel.Workbook;
                    if (wb.Name == wbName) {
                        wb.Activate();
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    UpdateTreeView();
                    return;
                }
            }

            if (e.Node.Level == 1) {
                // This is a sheet level node.
                string sheetName = e.Node.Text;
                var found = false;
                foreach (var o in Globals.ThisAddIn.Application.Workbooks[wbName].Sheets) {
                    var sheet = o as Microsoft.Office.Interop.Excel.Worksheet;
                    if (sheet.Name == sheetName) {
                        sheet.Activate();
                        found = true;
                    }
                }
                if (!found) {
                    UpdateTreeView();
                    return;
                }
            }
        }
    }
}
