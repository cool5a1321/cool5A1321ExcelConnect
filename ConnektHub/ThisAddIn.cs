using Autofac;
using Prospecta.ConnektHub.Models;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prospecta.ConnektHub
{
    public partial class ThisAddIn
    {
        #region Global Variables
        public Excel.Application excelApplication;
        #endregion
        #region Private Methods
        /// <summary>
        /// Method used to get the process id of the excel application
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lpdwProcessId"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        /// <summary>
        /// StartUp method of the add in
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            excelApplication = this.Application;
            var builder = new ContainerBuilder();
            Bootstrap.RegisterTypes(ref builder);
            GlobalMembers.InstanceGlobalMembers.Container = builder.Build();
        }
        /// <summary>
        /// Shutdown method of the addin
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (excelApplication != null)
            {
                int hWnd = excelApplication.Hwnd;
                uint processID;
                GetWindowThreadProcessId((IntPtr)hWnd, out processID);
                Process[] procs = Process.GetProcessesByName("EXCEL");
                foreach (Process p in procs)
                {
                    string activeWorkbookName = excelApplication.ActiveWorkbook.Name;
                    if (p.Id == processID && activeWorkbookName == "ConnektHubMetaData.xlsx")
                        p.Kill();
                }
                Marshal.FinalReleaseComObject(excelApplication);
            }
        }
        #endregion
        #region VSTO generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor....
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}