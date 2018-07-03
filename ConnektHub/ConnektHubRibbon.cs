using Autofac;
using ConnektHub.Models;
using Microsoft.Office.Tools.Ribbon;
using Prospecta.ConnektHub.Controllers;
using Prospecta.ConnektHub.Forms;
using Prospecta.ConnektHub.Helpers;
using Prospecta.ConnektHub.Models;
using Prospecta.ConnektHub.Services.HttpService;
using Prospecta.ConnektHub.Services.Modules;
using Prospecta.ConnektHub.Services.Translation;
using Prospecta.ConnektHub.Services.User;
using Prospecta.ConnektHub.SQLiteHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prospecta.ConnektHub
{
    public partial class ConnektHubRibbon
    {
        string userId = string.Empty, moduleId = string.Empty;
        Dictionary<string, string> dictModules = null;
        ProgressBarForm alert;
        Excel.Workbook excelWorkbook;
        [DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(int hWnd);
        private void XlBringToFront()
        { SetForegroundWindow(Globals.ThisAddIn.Application.Hwnd); }
        private bool IsExcelInteractive()
        {
            try
            {
                Globals.ThisAddIn.Application.Interactive = Globals.ThisAddIn.Application.Interactive;
                return true;
            }
            catch
            { return false; }
        }
        private void ExitEditMode()
        {
            if (!IsExcelInteractive())
            {
                Microsoft.Office.Interop.Excel.Range r = Globals.ThisAddIn.Application.ActiveCell;
                XlBringToFront();
                Globals.ThisAddIn.Application.ActiveWindow.Activate();
                SendKeys.Flush();
                SendKeys.SendWait("{ENTER}");
                r.Select();
            }
        }
        /// <summary>
        /// Ribbon Load Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConnektHubRibbon_Load(object sender, RibbonUIEventArgs e)
        { RibbonLoad(); }

        /// <summary>
        /// Login button event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLogin_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonLoad();

            IHttpRequest httpRequest = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IHttpRequest>();
            IUserService _login = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IUserService>(new TypedParameter(typeof(IHttpRequest), httpRequest));
            var loginForm = new LoginForm(_login);
            DialogResult dialogResult = loginForm.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                var result = loginForm.AuthenticationResult;
                if (result == 1)
                {
                    var userFullName = loginForm.StrFullName;
                    userId = loginForm.StrUserName;
                    btnUserDetails.Label = userFullName;
                    btnUserDetails.Visible = true;
                    btnUserDetails.Enabled = false;
                    btnModules.Enabled = true;
                    btnLogin.Enabled = false;
                }
            }
        }

        private void ExcelWorkbook_SheetFollowHyperlink(object Sh, Excel.Hyperlink Target)
        {
            try
            {
                if ("Drop Down Values".Equals(Target.Range.Text))
                {
                    Excel._Worksheet ws = (Excel._Worksheet)Sh;
                    int rowNumber = Target.Range.Row;
                    int fnColIndex = ExcelHelper.ConvertColumnNameToColumnIndex("FieldName", ws);
                    Excel.Range r = ws.Cells[rowNumber, fnColIndex] as Excel.Range;
                    Excel.Workbook workbook = (Excel.Workbook)ws.Parent;
                    Excel.Sheets oSheets = workbook.Worksheets;
                    Excel._Worksheet oSheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.DropDownSheetName, oSheets);

                    int targetSheetRow = oSheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                    targetSheetRow++;
                    fnColIndex = ExcelHelper.ConvertColumnNameToColumnIndex("Field Id", oSheet);
                    oSheet.Cells[targetSheetRow, fnColIndex] = r.Text;
                }
            }
            catch (Exception ex)
            { throw (ex); }
        }

        private void BtnModules_Click(object sender, RibbonControlEventArgs e)
        {
            IHttpRequest httpRequest = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IHttpRequest>();
            IModuleService moduleService = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IModuleService>(new TypedParameter(typeof(IHttpRequest), httpRequest));
            ModuleController moduleController = new ModuleController(moduleService);
            dictModules = moduleController.GetModulesList(userId);

            try
            {
                foreach (KeyValuePair<string, string> kv in dictModules)
                {
                    if (kv.Value != null)
                    {
                        RibbonButton moduleButton = Factory.CreateRibbonButton();
                        moduleButton.Name = "btn_" + kv.Key;
                        moduleButton.Label = kv.Value.ToString();
                        moduleButton.Click += OnModuleButton_Click;
                        dynamicMenuModules.Items.Add(moduleButton);
                    }
                }
                dynamicMenuModules.Visible = true;
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        private void OnModuleButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ExcelHelper.ToggleExcelEvents(Globals.ThisAddIn.excelApplication, false);
                RibbonButton rb = (RibbonButton)sender;
                RibbonMenu rm = (RibbonMenu)rb.Parent;
                rm.Label = rb.Label;
                dynamicMenuAdd.Enabled = true;
                foreach (RibbonButton addButton in dynamicMenuAdd.Items)
                {
                    addButton.Click += OnAddButton_Click;
                }
                ExcelHelper.ToggleExcelEvents(Globals.ThisAddIn.excelApplication, true);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        private void BtnImport_Click(object sender, RibbonControlEventArgs e)
        {
            Excel._Worksheet worksheet = Globals.ThisAddIn.excelApplication.ActiveSheet;
            Excel.Range oRange = worksheet.UsedRange;
            oRange.ClearContents();
            string moduleId = ExcelHelper.GetObjectIdFromText(dynamicMenuModules);
            IHttpRequest httpRequest = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IHttpRequest>();
            ITranslationService _translationService = GlobalMembers.InstanceGlobalMembers.Container.Resolve<ITranslationService>(new TypedParameter(typeof(IHttpRequest), httpRequest));
            TranslationController translationController = new TranslationController(_translationService);
            ExcelHelper.PopulateTranslationFields(worksheet, moduleId, translationController);
        }

        private void BtnExport_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper.ToggleExcelEvents(Globals.ThisAddIn.excelApplication, false);
            Cursor.Current = Cursors.WaitCursor;
            RibbonMenu ribbonMenu = (RibbonMenu)dynamicMenuAdd;
            string addMenuLabel = ribbonMenu.Label;
            InitializeVariables.InitializeDatabaseTables();
            if (backgroundWorker1.IsBusy != true)
            {
                alert = new ProgressBarForm();
                switch (addMenuLabel)
                {
                    case "Fields":
                        alert.Text = "Validating and Exporting Fields Data";
                        break;
                    case "Dropdowns":
                        alert.Text = "Exporting dropdown data";
                        break;
                    case "Descriptions":
                        alert.Text = "Exporting translated descriptions";
                        break;
                    case "Fields and Dropdowns":
                        alert.Text = "Validating Fields and Exporting Fields and Dropdowns data";
                        break;
                }
                alert.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            Cursor.Current = Cursors.Default;
            ExcelHelper.ToggleExcelEvents(Globals.ThisAddIn.excelApplication, true);
        }

        private void BackgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            string addMenuLabel = string.Empty;
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                bool isExported = false;
                ExitEditMode();
                RibbonMenu ribbonMenu = (RibbonMenu)dynamicMenuAdd;
                addMenuLabel = ribbonMenu.Label;
                IHttpRequest httpRequest = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IHttpRequest>();
                RibbonController ribbonController = new RibbonController(httpRequest);
                Excel._Worksheet worksheet = Globals.ThisAddIn.excelApplication.ActiveSheet;
                int colCount = worksheet.UsedRange.Columns.Count;
                int rowCount = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                string moduleId = ExcelHelper.GetObjectIdFromText(dynamicMenuModules);
                Dictionary<string, string> dictLanguage = GlobalMembers.InstanceGlobalMembers.DictionaryLanguageType;
                switch (addMenuLabel)
                {
                    case "Fields":
                        {
                            ExcelHelper.ClearLogColumnsForAllTheRows(worksheet);
                            for (int i = 4; i <= rowCount; i++)
                            {
                                if (worker.CancellationPending == true)
                                {
                                    e.Cancel = true;
                                    break;
                                }
                                else
                                {
                                    ribbonController.ValidateExcelData(i, colCount, worksheet);
                                    ribbonController.BuildAndExportFieldsJson(i, colCount, worksheet, moduleId, userId);
                                }
                                worker.ReportProgress((i * 100) / rowCount);
                            }
                            break;
                        }
                    case "Descriptions":
                        {
                            for (int i = 3; i <= rowCount; i++)
                            {
                                if (worker.CancellationPending == true)
                                {
                                    e.Cancel = true;
                                    break;
                                }
                                else
                                {
                                    if (!isExported)
                                    { ribbonController.BuildTranlatedDataJsonAndExport(worksheet, dictLanguage); }
                                    isExported = true;
                                    worker.ReportProgress((i * 100) / rowCount);
                                }
                            }
                            break;
                        }
                    case "Dropdowns":
                        {
                            for (int i = 3; i <= rowCount; i++)
                            {
                                if (worker.CancellationPending == true)
                                {
                                    e.Cancel = true;
                                    break;
                                }
                                else
                                {
                                    if (!isExported)
                                    { ribbonController.BuildDropDownJsonAndExport(worksheet, dictLanguage); }
                                    isExported = true;
                                    worker.ReportProgress((i * 100) / rowCount);
                                }
                            }
                            break;
                        }
                    case "Fields And Dropdowns":
                        {
                            Excel.Workbook workbook = GlobalMembers.InstanceGlobalMembers.ExcelApplication.ActiveWorkbook;
                            Excel.Sheets sheets = workbook.Sheets;
                            Excel._Worksheet dropDownWorkSheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.DropDownSheetName, sheets); ;
                            for (int i = 4; i <= rowCount; i++)
                            {
                                if (worker.CancellationPending == true)
                                {
                                    e.Cancel = true;
                                    break;
                                }
                                else
                                {
                                    ribbonController.ValidateExcelData(i, colCount, worksheet);
                                    ribbonController.BuildAndExportFieldsJson(i, colCount, worksheet, moduleId, userId);
                                    if (i == rowCount)
                                    { ribbonController.BuildDropDownJsonAndExport(dropDownWorkSheet, dictLanguage); }
                                }
                                worker.ReportProgress((i * 100) / rowCount);
                            }
                            break;
                        }
                }
            }
            catch (Exception ex)
            { throw (ex); }
        }

        private void BackgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            alert.Message = "In progress, please wait... ";
            alert.ProgressValue = e.ProgressPercentage;
            System.Threading.Thread.Sleep(20);
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        { alert.Close(); }

        private void OnAddButton_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton rb = (RibbonButton)sender;
            RibbonMenu rm = (RibbonMenu)rb.Parent;
            string addMenuLabel = string.Empty;
            addMenuLabel = rb.Label;
            rm.Label = rb.Label;
            Excel.Sheets sheets = Globals.ThisAddIn.excelApplication.Sheets;
            Excel._Worksheet worksheet;
            Excel.Range oRange;
            ExcelHelper.ToggleExcelEvents(Globals.ThisAddIn.excelApplication, false);
            switch (addMenuLabel)
            {
                case "Fields":
                    {
                        if (btnImport.Enabled)
                        { btnImport.Enabled = false; }
                        worksheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.MetaDataSheetName, sheets);
                        oRange = worksheet.UsedRange;
                        oRange.ClearContents();
                        worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        ExcelHelper.DisplayMetaDataInformationOnSheet(worksheet, dictModules, addMenuLabel);
                        break;
                    }
                case "Descriptions":
                    {
                        if (!btnImport.Enabled)
                        { btnImport.Enabled = true; }
                        worksheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.TranslationSheetName, sheets);
                        oRange = worksheet.UsedRange;
                        oRange.ClearContents();
                        ExcelHelper.PopulateTranslationHeader(worksheet);
                        worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        worksheet.Activate();
                        break;
                    }
                case "Dropdowns":
                    {
                        if (btnImport.Enabled)
                        { btnImport.Enabled = false; }
                        worksheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.DropDownSheetName, sheets);
                        oRange = worksheet.UsedRange;
                        oRange.ClearContents();
                        ExcelHelper.PopulateDropDownValuesHeader(worksheet);
                        worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        worksheet.Activate();
                        break;
                    }
                case "Fields And Dropdowns":
                    {
                        if (btnImport.Enabled)
                        { btnImport.Enabled = false; }
                        worksheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.MetaDataSheetName, sheets);
                        oRange = worksheet.UsedRange;
                        oRange.ClearContents();
                        Excel._Worksheet worksheet1 = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.DropDownSheetName, sheets);
                        oRange = worksheet1.UsedRange;
                        oRange.ClearContents();
                        ExcelHelper.PopulateDropDownValuesHeader(worksheet1);
                        worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        ExcelHelper.DisplayMetaDataInformationOnSheet(worksheet, dictModules, addMenuLabel);
                        worksheet1.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        worksheet.Activate();
                        break;
                    }
            }

            if (!btnExport.Enabled)
            { btnExport.Enabled = true; }

            ExcelHelper.HideExcelSheetsExceptTheSelectedOne(rm.Label.ToString(), sheets);
            ExcelHelper.ToggleExcelEvents(Globals.ThisAddIn.excelApplication, true);
        }

        private void RibbonLoad()
        {
           try
            {
                if (Globals.ThisAddIn.excelApplication != null)
                {
                    string directoryPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string filePath = Path.Combine(directoryPath, "MasterDataOnline.db");
                    GlobalMembers.InstanceGlobalMembers.SqliteDatabase = new SQLiteDatabase(filePath);
                    InitializeGlobalVariables();
                    InitializeVariables.InitializeDatabaseTables();
                    InitializeVariables.InitializeClassVariables();

                    this.excelWorkbook = Globals.ThisAddIn.excelApplication.ActiveWorkbook;

                    ExcelHelper.ReadTheWebServicePathFromConfigSheet(excelWorkbook);

                    Globals.ThisAddIn.excelApplication.WorkbookBeforeClose -= ExcelApplication_WorkbookBeforeClose;
                    this.excelWorkbook.SheetChange -= ExcelWorkbook_SheetChange;
                    this.excelWorkbook.SheetFollowHyperlink -= ExcelWorkbook_SheetFollowHyperlink;

                    Globals.ThisAddIn.excelApplication.WorkbookBeforeClose += ExcelApplication_WorkbookBeforeClose;
                    this.excelWorkbook.SheetChange += ExcelWorkbook_SheetChange;
                    this.excelWorkbook.SheetFollowHyperlink += ExcelWorkbook_SheetFollowHyperlink;
                }
                btnUserDetails.Visible = false;
                btnModules.Enabled = false;
                btnImport.Enabled = false;
                btnExport.Enabled = false;
                dynamicMenuModules.Visible = false;
                dynamicMenuAdd.Enabled = false;
            }
            catch (Exception ex)
            { throw (ex); }
        }
        
        private void ExcelWorkbook_SheetChange(object Sh, Excel.Range Target)
        {
            try
            {
                Excel._Worksheet ws = (Excel._Worksheet)Sh;
                switch (ws.Name)
                {
                    case "Introduction":
                    case "Configuration":
                    case "DropDownValues":
                    case "Translation":
                        return;
                }

                if (Target.Cells.Count > 1)
                {
                    foreach (Excel.Range cell in Target.Cells)
                    {
                        ExcelHelper.ChangeFieldNameToUpperCase(ws, cell);
                        ExcelHelper.RemoveTargetBorder(ws, cell);
                        ExcelHelper.CheckDependencies(ws, cell);
                        ExcelHelper.PopulateDefaultFieldLength(ws, cell);
                        ExcelHelper.AddHyperlinkToACell(ws, cell, dynamicMenuAdd);
                    }
                }
                else
                {
                    ExcelHelper.ChangeFieldNameToUpperCase(ws, Target);
                    ExcelHelper.RemoveTargetBorder(ws, Target);
                    ExcelHelper.CheckDependencies(ws, Target);
                    ExcelHelper.PopulateDefaultFieldLength(ws, Target);
                    ExcelHelper.AddHyperlinkToACell(ws, Target, dynamicMenuAdd);
                }
            }
            catch (Exception ex)
            { throw (ex); }

        }

        private void ExcelApplication_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            Excel.Sheets excelSheets = Wb.Worksheets;
            Excel.Worksheet excelWorksheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.MetaDataSheetName, excelSheets);
            Excel.Range excelRange = excelWorksheet.UsedRange;
            excelRange.ClearContents();
            excelRange.ClearFormats();
            excelRange.ClearComments();
        }

        private void ExcelApplication_WorkbookOpen(Excel.Workbook Wb)
        { ExcelHelper.ReadTheWebServicePathFromConfigSheet(Wb); }

        /// <summary>
        /// Method used to initialize the global variables
        /// </summary>
        private void InitializeGlobalVariables()
        {
            GlobalMembers.InstanceGlobalMembers.DtMetadataData = new System.Data.DataTable();
            GlobalMembers.InstanceGlobalMembers.DtFieldTypeValues = new System.Data.DataTable();
            GlobalMembers.InstanceGlobalMembers.DtDataTypeValues = new System.Data.DataTable();
            GlobalMembers.InstanceGlobalMembers.DtAttachmentTypeValues = new System.Data.DataTable();
            GlobalMembers.InstanceGlobalMembers.DtTrueFalseValues = new System.Data.DataTable();
            GlobalMembers.InstanceGlobalMembers.DtYesNoValues = new System.Data.DataTable();

            GlobalMembers.InstanceGlobalMembers.IntroductionSheetName = "Introduction";
            GlobalMembers.InstanceGlobalMembers.ConfigurationSheetName = "Configuration";
            GlobalMembers.InstanceGlobalMembers.MetaDataSheetName = "MetaData";
            GlobalMembers.InstanceGlobalMembers.DropDownSheetName = "DropDownValues";
            GlobalMembers.InstanceGlobalMembers.TranslationSheetName = "Translation";


            GlobalMembers.InstanceGlobalMembers.ListMetadataData = new List<FieldMetaDataModel>();
            GlobalMembers.InstanceGlobalMembers.ListTranslationHeaders = new List<TranslationHeader>();
            GlobalMembers.InstanceGlobalMembers.ListDropdownHeadersModels = new List<DropdownHeadersModel>();
            GlobalMembers.InstanceGlobalMembers.DictionaryModulesList = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryDataTypes = new Dictionary<string, string>(); ;
            GlobalMembers.InstanceGlobalMembers.DictionaryFieldTypes = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryAttachmentFileTypes = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryYesNo = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryTrueFalse = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryLanguageType = new Dictionary<string, string>();
            GlobalMembers.InstanceGlobalMembers.DictionaryDefaultFieldLength = new Dictionary<string, int>();
            GlobalMembers.InstanceGlobalMembers.ListDropdownHeaders = new List<string>();
            GlobalMembers.InstanceGlobalMembers.ListSqlKeyWords = new List<string>();
        }
    }
}