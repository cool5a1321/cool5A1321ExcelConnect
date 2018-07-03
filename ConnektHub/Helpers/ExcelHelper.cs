using Autofac;
using ConnektHub.Models;
using Microsoft.Office.Tools.Ribbon;
using Prospecta.ConnektHub.Controllers;
using Prospecta.ConnektHub.Models;
using Prospecta.ConnektHub.Services.HttpService;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prospecta.ConnektHub.Helpers
{
    public class ExcelHelper
    {
        #region Public Methods
        /// <summary>
        /// Method used to toggle excel events
        /// </summary>
        /// <param name="oXL"></param>
        /// <param name="boolVal"></param>
        public static void ToggleExcelEvents(Excel.Application oXL, bool boolVal)
        {
            oXL.ScreenUpdating = boolVal;
            oXL.DisplayAlerts = boolVal;
            oXL.EnableEvents = boolVal;
        }
        /// <summary>
        /// Get the object id from text
        /// </summary>
        /// <param name="ribbonMenu"></param>
        /// <returns></returns>
        public static string GetObjectIdFromText(RibbonMenu ribbonMenu)
        {
            string objectType = string.Empty;
            try
            {
                for (int i = 0; i < ribbonMenu.Items.Count; i++)
                {
                    RibbonButton ribbonButton = (RibbonButton)ribbonMenu.Items[i];
                    if (ribbonButton.Label == ribbonMenu.Label)
                    { objectType = ribbonButton.Name.Split('_')[1]; break; }
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            return objectType;
        }
        /// <summary>
        /// Method to populate translation fields on excel sheet
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="moduleId"></param>
        /// <param name="translationController"></param>
        public static void PopulateTranslationFields(Excel._Worksheet sh, string moduleId, TranslationController translationController)
        {
            Excel.Range oRange;
            try
            {
                PopulateTranslationHeader(sh);
                int colCount = 1, minColIndex = 0, maxColIndex = 0, rowCount = 0;
                List<TranslationData> translationData = translationController.GetFieldIdNDescriptionInEnglish(moduleId);
                int targetSheetRow = sh.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                rowCount = targetSheetRow;
                targetSheetRow++;
                foreach (var item in translationData)
                {
                    sh.Cells[targetSheetRow, colCount] = item.FieldId;
                    sh.Cells[targetSheetRow, colCount + 1] = item.FieldDescription;
                    targetSheetRow++;
                }
                minColIndex = colCount;
                maxColIndex = sh.UsedRange.Columns.Count;

                oRange = sh.UsedRange;

                oRange.ClearFormats();
                oRange = sh.Cells[rowCount - 1, maxColIndex] as Excel.Range;
                oRange.EntireRow.Hidden = true;

                oRange = sh.Range[sh.Cells[1, minColIndex], sh.Cells[targetSheetRow - 1, maxColIndex]] as Excel.Range;
                oRange.EntireColumn.AutoFit();

                oRange = sh.Range[sh.Cells[rowCount, 1], sh.Cells[rowCount, maxColIndex]];
                oRange.Interior.Color = ColorTranslator.ToOle(Color.DarkGray);
                oRange.Cells.Font.Bold = true;
                oRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to populate the translation header in excel sheet
        /// </summary>
        /// <param name="sh"></param>
        public static void PopulateTranslationHeader(Excel._Worksheet sh)
        {
            int rowCount = 0, colCount = 0, maxColIndex = 0, minColIndex = 0;
            Excel.Range oRange;
            try
            {
                oRange = sh.UsedRange;
                oRange.ClearContents();
                List<TranslationHeader> lstTranslationHeaders = GlobalMembers.InstanceGlobalMembers.ListTranslationHeaders;

                rowCount = 2;
                colCount = 1;
                minColIndex = 1;
                maxColIndex = lstTranslationHeaders.Count;
                Dictionary<string, string> dictLanguageTypes = new Dictionary<string, string>();
                foreach (KeyValuePair<string, string> kv in GlobalMembers.InstanceGlobalMembers.DictionaryLanguageType)
                {
                    if (!kv.Key.Equals("English"))
                    {
                        dictLanguageTypes.Add(kv.Key, kv.Value);
                    }
                }


                foreach (var item in lstTranslationHeaders)
                {
                    oRange = sh.Cells[rowCount + 1, colCount] as Excel.Range;
                    sh.Cells[rowCount - 1, colCount] = item.FieldName;
                    sh.Cells[rowCount, colCount] = item.Description;
                    if (item.IsDropDown)
                    { CreateDropDown(dictLanguageTypes, oRange); }
                    colCount++;
                }

                oRange = sh.Cells[rowCount - 1, maxColIndex] as Excel.Range;
                oRange.EntireRow.Hidden = true;

                oRange = sh.Range[sh.Cells[1, 1], sh.Cells[2, maxColIndex]] as Excel.Range;
                oRange.EntireColumn.AutoFit();

                oRange = sh.Range[sh.Cells[rowCount, 1], sh.Cells[rowCount, maxColIndex]];
                oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                oRange.Cells.Font.Bold = true;
                oRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                var sRange = sh.Range[sh.Cells[rowCount + 1, minColIndex], sh.Cells[rowCount + 1, maxColIndex]].EntireRow;
                var dRange = sh.Range[sh.Cells[rowCount + 2, minColIndex], sh.Cells[1000, maxColIndex]].EntireRow;
                sRange.Copy(dRange);
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to dispaly metadata data on screen
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="dictModulesList"></param>
        /// <param name="addMenuItem"></param>
        public static void DisplayMetaDataInformationOnSheet(Excel._Worksheet sh, Dictionary<string, string> dictModulesList, string addMenuItem)
        {
            int maxColumnIndex = 0, startRowIndex = 3, colCount = 2, startColumnIndex = 0, minColumnIndex = 1;

            Excel.Range oRange;
            try
            {
                sh.Activate();

                if (GlobalMembers.InstanceGlobalMembers.DtMetadataData.Rows.Count == 0)
                { GlobalMembers.InstanceGlobalMembers.DtMetadataData = PopulateMetadataData(); }

                maxColumnIndex = GlobalMembers.InstanceGlobalMembers.DtMetadataData.Rows.Count;

                if (addMenuItem.Equals("Fields And Dropdowns"))
                { maxColumnIndex += 2; }
                else
                { maxColumnIndex += 1; }

                sh.Cells[startRowIndex - 1, 1] = "Log";
                sh.Cells[startRowIndex, 1] = "Log";

                ((Excel.Range)sh.Cells[startRowIndex - 2, 1]).Interior.Color = ColorTranslator.ToOle(Color.DarkGray);
                ((Excel.Range)sh.Cells[startRowIndex, 1]).Interior.Color = ColorTranslator.ToOle(Color.DarkGray);

                DataView dataView = new DataView(GlobalMembers.InstanceGlobalMembers.DtMetadataData);
                DataTable dtDistinctValue = dataView.ToTable(true, new string[] { "TabName", "Colour" });
                foreach (DataRow dr in dtDistinctValue.Rows)
                {
                    startColumnIndex = colCount;
                    string filterCondition = string.Format("TabName = '{0}'", new string[] { dr["TabName"].ToString() });
                    DataRow[] dataRowArray = GlobalMembers.InstanceGlobalMembers.DtMetadataData.Select(filterCondition);
                    foreach (DataRow row in dataRowArray)
                    {
                        oRange = sh.Cells[startRowIndex, colCount] as Excel.Range;
                        oRange.ClearComments();
                        sh.Cells[startRowIndex - 2, colCount] = dr["TabName"].ToString();
                        sh.Cells[startRowIndex - 1, colCount] = row["FieldName"].ToString();
                        sh.Cells[startRowIndex, colCount] = row["Description"].ToString();
                        /* to show a star and change the color of the star to red if the field is mandatory */
                        if (row["Mandatory"].ToString().Equals("1"))
                        {
                            sh.Cells[startRowIndex, colCount] = row["Description"].ToString() + " *";
                            string fieldDescription = oRange.Text;
                            oRange.Characters[fieldDescription.LastIndexOf(' ') + 1, fieldDescription.Length].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                        /* to show a star and change the color of the star to red if the field is mandatory */
                        if (!string.IsNullOrEmpty(row["HelpText"].ToString()))
                        { oRange.AddComment(row["HelpText"].ToString()); }

                        if (row["IsDropDown"].ToString().Equals("1"))
                        {
                            oRange = sh.Cells[startRowIndex + 1, colCount] as Excel.Range;
                            if (row["DropDownType"].ToString().Equals("TrueFalse"))
                            {
                                if (GlobalMembers.InstanceGlobalMembers.DtTrueFalseValues.Rows.Count == 0)
                                { PopulateDropDowns("TrueFalseValues"); }
                                CreateDropDown(GlobalMembers.InstanceGlobalMembers.DtTrueFalseValues, oRange);
                            }
                            else if (row["DropDownType"].ToString().Equals("YesNo"))
                            {
                                if (GlobalMembers.InstanceGlobalMembers.DtYesNoValues.Rows.Count == 0)
                                { PopulateDropDowns("YesNoValues"); }
                                CreateDropDown(GlobalMembers.InstanceGlobalMembers.DtYesNoValues, oRange);
                            }
                            else if (row["DropDownType"].ToString().Equals("Other"))
                            {
                                if (row["FieldName"].ToString().Equals("FieldType"))
                                {
                                    if (GlobalMembers.InstanceGlobalMembers.DtFieldTypeValues.Rows.Count == 0)
                                    { PopulateDropDowns("FieldTypeValues"); }
                                    CreateDropDown(GlobalMembers.InstanceGlobalMembers.DtFieldTypeValues, oRange);
                                }
                                else if (row["FieldName"].ToString().Equals("DataType"))
                                {
                                    if (GlobalMembers.InstanceGlobalMembers.DtDataTypeValues.Rows.Count == 0)
                                    { PopulateDropDowns("DataTypeValues"); }
                                    CreateDropDown(GlobalMembers.InstanceGlobalMembers.DtDataTypeValues, oRange);
                                }
                                else if (row["FieldName"].ToString().Equals("refObjId"))
                                { CreateDropDown(dictModulesList, oRange); }
                                else if (row["FieldName"].ToString().Equals("AttachmentFileType"))
                                {
                                    if (GlobalMembers.InstanceGlobalMembers.DtAttachmentTypeValues.Rows.Count == 0)
                                    { PopulateDropDowns("AttachmentTypeValues"); }
                                    CreateDropDown(GlobalMembers.InstanceGlobalMembers.DtAttachmentTypeValues, oRange);
                                }
                            }
                        }
                        colCount++;
                    }
                    var mergeRange = sh.Range[sh.Cells[1, startColumnIndex], sh.Cells[1, colCount - 1]];
                    mergeRange.Merge(true);
                    Color c = Color.FromName(dr["Colour"].ToString());
                    mergeRange.Interior.Color = ColorTranslator.ToOle(c);
                    mergeRange.Cells.Font.Bold = true;
                    mergeRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    mergeRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                if (addMenuItem.Equals("Fields And Dropdowns"))
                {
                    sh.Cells[startRowIndex - 1, maxColumnIndex] = "DropDownValues";
                    sh.Cells[startRowIndex, maxColumnIndex] = "Drop Down Values";
                    ((Excel.Range)sh.Cells[startRowIndex - 2, maxColumnIndex]).Interior.Color = ColorTranslator.ToOle(Color.DarkGray);
                    ((Excel.Range)sh.Cells[startRowIndex, maxColumnIndex]).Interior.Color = ColorTranslator.ToOle(Color.DarkGray);
                }



                oRange = sh.Range[sh.Cells[startRowIndex - 2, minColumnIndex], sh.Cells[startRowIndex, maxColumnIndex]] as Excel.Range;
                oRange.EntireColumn.AutoFit();

                oRange = sh.Range[sh.Cells[startRowIndex - 2, minColumnIndex], sh.Cells[startRowIndex, maxColumnIndex]] as Excel.Range;
                oRange.RowHeight = 20;


                oRange = sh.Range[sh.Cells[startRowIndex, minColumnIndex], sh.Cells[startRowIndex, maxColumnIndex]] as Excel.Range;
                oRange.RowHeight = 20;
                oRange.Interior.Color = ColorTranslator.ToOle(Color.DarkGray);
                oRange.Cells.Font.Bold = true;
                oRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                oRange = sh.Cells[startRowIndex - 1, maxColumnIndex] as Excel.Range;
                oRange.EntireRow.Hidden = true;

                var sRange = sh.Range[sh.Cells[startRowIndex + 1, minColumnIndex], sh.Cells[startRowIndex + 1, maxColumnIndex]].EntireRow;
                var dRange = sh.Range[sh.Cells[startRowIndex + 2, minColumnIndex], sh.Cells[1000, maxColumnIndex]].EntireRow;
                sRange.Copy(dRange);

                var cellRange = sh.Cells[startRowIndex - 2, minColumnIndex] as Excel.Range;
                cellRange.Select();
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to populate the dropdown values from header
        /// </summary>
        /// <param name="sh"></param>
        public static void PopulateDropDownValuesHeader(Excel._Worksheet sh)
        {
            Excel.Range oRange;
            try
            {
                List<DropdownHeadersModel> lstDDHeaderModels = GlobalMembers.InstanceGlobalMembers.ListDropdownHeadersModels;
                Dictionary<string, string> dictLanguageType = GlobalMembers.InstanceGlobalMembers.DictionaryLanguageType;
                int rowCount = 2, colCount = 1;
                foreach (var item in lstDDHeaderModels)
                {
                    sh.Cells[rowCount - 1, colCount] = item.FieldName;
                    sh.Cells[rowCount, colCount] = item.Description;
                    Excel.Range dropDownRange = sh.Cells[rowCount + 1, colCount] as Excel.Range;
                    if (item.IsDropDown)
                    { CreateDropDown(dictLanguageType, dropDownRange); }
                    colCount++;
                }
                oRange = sh.Cells[rowCount - 1, colCount - 1] as Excel.Range;
                oRange.EntireRow.Hidden = true;

                oRange = sh.Range[sh.Cells[1, 1], sh.Cells[2, colCount - 1]] as Excel.Range;
                oRange.EntireColumn.AutoFit();

                oRange = sh.Range[sh.Cells[rowCount, 1], sh.Cells[rowCount, colCount - 1]];
                oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                oRange.Cells.Font.Bold = true;
                oRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                var sRange = sh.Range[sh.Cells[rowCount + 1, 1], sh.Cells[rowCount + 1, colCount - 1]].EntireRow;
                var dRange = sh.Range[sh.Cells[rowCount + 2, 1], sh.Cells[1000, colCount - 1]].EntireRow;
                sRange.Copy(dRange);
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to create dropdown fetching data from dictionary
        /// </summary>
        /// <param name="dict"></param>
        /// <param name="cellRange"></param>
        public static void CreateDropDown(Dictionary<string, string> dict, Excel.Range cellRange)
        {
            var itemArray = dict.Keys.ToArray();
            var dropList = string.Join(",", itemArray);
            cellRange.Validation.Delete();
            cellRange.Validation.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertInformation,
                Excel.XlFormatConditionOperator.xlBetween,
                dropList,
                Type.Missing
            );
            cellRange.Validation.IgnoreBlank = true;
            cellRange.Validation.InCellDropdown = true;
        }
        /// <summary>
        /// Method used to create dropdown fetching the data from sqlite
        /// </summary>
        /// <param name="dtDropDown"></param>
        /// <param name="cellRange"></param>
        private static void CreateDropDown(DataTable dtDropDown, Excel.Range cellRange)
        {
            var itemArray = dtDropDown.AsEnumerable().Select(x => x.Field<string>("Code")).ToList();
            var dropList = string.Join(",", itemArray);

            cellRange.Validation.Delete();
            cellRange.Validation.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertInformation,
                Excel.XlFormatConditionOperator.xlBetween,
                dropList,
                Type.Missing
            );
            cellRange.Validation.IgnoreBlank = true;
            cellRange.Validation.InCellDropdown = true;
        }
        /// <summary>
        /// Hide sheets based on the menuItem selected from dropdown
        /// </summary>
        /// <param name="menuItem"></param>
        /// <param name="sheets"></param>
        public static void HideExcelSheetsExceptTheSelectedOne(string menuItem, Excel.Sheets sheets)
        {
            foreach (Excel.Worksheet sh in sheets)
            {
                if (!sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.IntroductionSheetName) && !sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.ConfigurationSheetName))
                {
                    switch (menuItem)
                    {
                        case "Fields":
                            {
                                if (!(sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.MetaDataSheetName)))
                                {
                                    if (sh.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                                    { sh.Visible = Excel.XlSheetVisibility.xlSheetHidden; }
                                }
                                break;
                            }
                        case "Descriptions":
                            {
                                if (!(sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.TranslationSheetName)))
                                {
                                    if (sh.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                                    { sh.Visible = Excel.XlSheetVisibility.xlSheetHidden; }
                                }
                                break;
                            }
                        case "Dropdowns":
                            {
                                if (!(sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.DropDownSheetName)))
                                {
                                    if (sh.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                                    { sh.Visible = Excel.XlSheetVisibility.xlSheetHidden; }
                                }
                                break;
                            }
                        case "Fields And Dropdowns":
                            {
                                if (!sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.DropDownSheetName) && !sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.MetaDataSheetName))
                                {
                                    if (sh.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                                    { sh.Visible = Excel.XlSheetVisibility.xlSheetHidden; }
                                }
                                break;
                            }
                    }
                }
            }
        }
        /// <summary>
        /// Method to check duplicate value in field name
        /// </summary>
        /// <param name="sh"></param>
        public static void DuplicacyCheckInFieldName(Excel._Worksheet sh)
        {
            int rowCount = 0;
            Excel.Range or, dr;
            rowCount = sh.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            for (int rowCnt = 4; rowCnt <= rowCount - 1; rowCnt++)
            {
                or = sh.Cells[rowCnt, 2] as Excel.Range;
                for (int rowCnt1 = rowCnt + 1; rowCnt1 <= rowCount; rowCnt1++)
                {
                    dr = sh.Cells[rowCnt1, 2] as Excel.Range;
                    if (!(string.IsNullOrEmpty(or.Text) && string.IsNullOrEmpty(dr.Text)))
                    {
                        if (or.Text == dr.Text)
                        {
                            Excel.Range ocr = sh.Cells[rowCnt, 1] as Excel.Range;
                            Excel.Range dcr = sh.Cells[rowCnt1, 1] as Excel.Range;
                            ocr.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                            dcr.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                            ocr.Value = "Duplicate value found in fieldname columns";
                            dcr.Value = "Duplicate value found in fieldname columns";
                            break;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Method used to convert column name to column index
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="sh"></param>
        /// <returns></returns>
        public static int ConvertColumnNameToColumnIndex(string columnName, Excel._Worksheet sh)
        {
            int columnIndex = 0, colCount = 0;
            try
            {
                colCount = sh.UsedRange.Columns.Count;
                for (int i = 1; i <= colCount; i++)
                {
                    var r = sh.Cells[2, i] as Excel.Range;
                    if (r.Text == columnName)
                    { columnIndex = i; break; }
                }
            }
            catch (Exception ex)
            { throw (ex); }
            return columnIndex;
        }
        /// <summary>
        /// Method used to retrieve the value of the dropdown from the code in the dropdown
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        public static string GetValueFromCodeForDropDown(DataTable dt, string code)
        {
            string resultValue = string.Empty, filterCondition = string.Empty;
            filterCondition = string.Format("Code = '{0}'", new string[] { code });
            DataRow[] drArray = dt.Select(filterCondition);
            if (drArray.Length > 0)
            { resultValue = drArray[0]["TEXT"].ToString(); }
            return resultValue;
        }
        /// <summary>
        /// Method used to populate the data from list to sqlite
        /// </summary>
        /// <returns></returns>
        public static DataTable PopulateMetadataData()
        {
            DataTable dt = new DataTable();
            string[] columnName = new string[] { "FieldName", "Description", "TabName", "HelpText", "IsDropDown", "DropDownType", "Colour", "Mandatory" };
            string[] columnDataTypes = new string[] { "TEXT", "TEXT", "TEXT", "TEXT", "INTEGER", "TEXT", "TEXT", "INTEGER" };
            if (!GlobalMembers.InstanceGlobalMembers.SqliteDatabase.CheckIfTableExists("MetadataData"))
            {
                GlobalMembers.InstanceGlobalMembers.SqliteDatabase.CreateTable("MetadataData", columnName, columnDataTypes);
            }
            dt = GlobalMembers.InstanceGlobalMembers.SqliteDatabase.GetDataTable("select * from MetadataData");
            if (dt.Rows.Count == 0)
            {
                var metaDataItems = GlobalMembers.InstanceGlobalMembers.ListMetadataData;
                var properties = typeof(FieldMetaDataModel).GetProperties();
                foreach (var item in metaDataItems)
                {
                    Dictionary<string, string> dictMetaDataItems = new Dictionary<string, string>();
                    foreach (var property in properties)
                    {

                        if (property.PropertyType.Equals(typeof(Boolean)))
                        {
                            Boolean b = (Boolean)property.GetValue(item, null);
                            string s = string.Empty;
                            s = b.Equals(false) ? "0" : "1";
                            dictMetaDataItems.Add(property.Name, s);
                        }
                        else if (property.PropertyType.Equals(typeof(Color)))
                        {
                            Color c = (Color)property.GetValue(item, null);
                            dictMetaDataItems.Add(property.Name, c.Name.ToString());
                        }
                        else
                        { dictMetaDataItems.Add(property.Name, (string)property.GetValue(item, null)); }
                    }
                    long val = GlobalMembers.InstanceGlobalMembers.SqliteDatabase.Insert("MetadataData", dictMetaDataItems);
                }
                dt = GlobalMembers.InstanceGlobalMembers.SqliteDatabase.GetDataTable("select * from MetadataData");
            }
            return dt;
        }
        /// <summary>
        /// Method used to populate drop down fetching the data from sqlite database
        /// </summary>
        /// <param name="tableName"></param>
        public static void PopulateDropDowns(string tableName)
        {
            DataTable dt = new DataTable();
            if (!GlobalMembers.InstanceGlobalMembers.SqliteDatabase.CheckIfTableExists(tableName))
            { GlobalMembers.InstanceGlobalMembers.SqliteDatabase.CreateTable(tableName, new string[] { "Code", "Text" }, new string[] { "TEXT", "TEXT" }); }
            dt = GlobalMembers.InstanceGlobalMembers.SqliteDatabase.GetDataTable("select * from " + tableName);
            if (dt.Rows.Count == 0)
            {
                long l = 0;
                Dictionary<string, string> dict = new Dictionary<string, string>();
                if (tableName.Equals("FieldTypeValues"))
                { dict = GlobalMembers.InstanceGlobalMembers.DictionaryFieldTypes; }
                else if (tableName.Equals("DataTypeValues"))
                { dict = GlobalMembers.InstanceGlobalMembers.DictionaryDataTypes; }
                else if (tableName.Equals("AttachmentTypeValues"))
                { dict = GlobalMembers.InstanceGlobalMembers.DictionaryAttachmentFileTypes; }
                else if (tableName.Equals("TrueFalseValues"))
                { dict = GlobalMembers.InstanceGlobalMembers.DictionaryTrueFalse; }
                else if (tableName.Equals("YesNoValues"))
                { dict = GlobalMembers.InstanceGlobalMembers.DictionaryYesNo; }

                foreach (KeyValuePair<string, string> kv in dict)
                {
                    Dictionary<string, string> finalDictionary = new Dictionary<string, string>
                    {
                        { "Code", kv.Key },
                        { "Text", kv.Value }
                    };
                    l = GlobalMembers.InstanceGlobalMembers.SqliteDatabase.Insert(tableName, finalDictionary);
                }
                dt = GlobalMembers.InstanceGlobalMembers.SqliteDatabase.GetDataTable("select * from " + tableName);
            }
            if (tableName.Equals("FieldTypeValues"))
            { GlobalMembers.InstanceGlobalMembers.DtFieldTypeValues = dt; }
            else if (tableName.Equals("DataTypeValues"))
            { GlobalMembers.InstanceGlobalMembers.DtDataTypeValues = dt; }
            else if (tableName.Equals("AttachmentTypeValues"))
            { GlobalMembers.InstanceGlobalMembers.DtAttachmentTypeValues = dt; }
            else if (tableName.Equals("TrueFalseValues"))
            { GlobalMembers.InstanceGlobalMembers.DtTrueFalseValues = dt; }
            else if (tableName.Equals("YesNoValues"))
            { GlobalMembers.InstanceGlobalMembers.DtYesNoValues = dt; }
        }
        /// <summary>
        /// reading the web service path from the config sheet
        /// </summary>
        /// <param name="workBook"></param>
        public static void ReadTheWebServicePathFromConfigSheet(Excel.Workbook workBook)
        {
            IHttpRequest httpRequest = GlobalMembers.InstanceGlobalMembers.Container.Resolve<IHttpRequest>();
            if (workBook.Name == "ConnektHubMetaData.xlsx")
            {
                Excel.Sheets oSheets = workBook.Worksheets;
                Excel._Worksheet oSheet = HelperUtil.GetSheetNameFromGroupOfSheets(GlobalMembers.InstanceGlobalMembers.ConfigurationSheetName, oSheets);
                Excel.Range oRange = oSheet.Cells[1, 2] as Excel.Range;
                httpRequest.BaseURL = oRange.Text;
            }

        }
        /// <summary>
        /// Method used to change the field name to upper case
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="range"></param>
        public static void ChangeFieldNameToUpperCase(Excel._Worksheet ws, Excel.Range range)
        {
            int fnColIndex = 0;
            string fieldNameValue = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(range.Text))
                {
                    fnColIndex = ConvertColumnNameToColumnIndex("FieldName", ws);
                    if (fnColIndex > 0)
                    {
                        Excel.Range fnRange = ws.Cells[range.Row, fnColIndex] as Excel.Range;
                        fieldNameValue = fnRange.Text;
                        if (fieldNameValue.Equals(range.Text))
                        {
                            if (!HelperUtil.IsAllUpper(fieldNameValue))
                            { range.Value = fieldNameValue.ToUpper(); }
                        }
                    }
                }
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to remove the border color of the cell
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="range"></param>
        public static void RemoveTargetBorder(Excel._Worksheet ws, Excel.Range range)
        {
            try
            {
                if (range.Borders.Color == 0x0000FF)
                { range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone; }
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to check the depdency and chaning the border color or adding a default value to the cell
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="range"></param>
        public static void CheckDependencies(Excel._Worksheet ws, Excel.Range range)
        {
            try
            {
                string[] arrayDependencies = null;
                string dependencyValue = string.Empty;
                int colIndex = 0;
                if (!string.IsNullOrEmpty(range.Text))
                {
                    if (GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies.ContainsKey(range.Text))
                    {
                        arrayDependencies = GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies[range.Text].Split(',');
                        for (int i = 0; i < arrayDependencies.Length; i++)
                        {
                            dependencyValue = arrayDependencies[i].Trim();
                            colIndex = ExcelHelper.ConvertColumnNameToColumnIndex(dependencyValue, ws);
                            Excel.Range r = ws.Cells[range.Row, colIndex] as Excel.Range;
                            if (dependencyValue.Equals("textAreaLength"))
                            { r.Value = "60"; }
                            else if (dependencyValue.Equals("textAreaWidth"))
                            { r.Value = "230"; }
                            else if (dependencyValue.Equals("optionLimit"))
                            { r.Value = "1"; }
                            else
                            { r.Borders.Color = ColorTranslator.ToOle(Color.Red); }
                        }
                    }
                }
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to check the field length
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="range"></param>
        public static void PopulateDefaultFieldLength(Excel._Worksheet ws, Excel.Range range)
        {
            int flColIndex = 0, ftColIndex = 0;
            try
            {
                if (!string.IsNullOrEmpty(range.Text))
                {
                    flColIndex = ConvertColumnNameToColumnIndex("FieldLength", ws);
                    if (!flColIndex.Equals(0))
                    {
                        ftColIndex = ConvertColumnNameToColumnIndex("FieldType", ws);
                        if (!ftColIndex.Equals(0))
                        {
                            Excel.Range ftRange = ws.Cells[range.Row, ftColIndex] as Excel.Range;
                            if (range.Text == ftRange.Text)
                            {
                                if (GlobalMembers.InstanceGlobalMembers.DictionaryDefaultFieldLength.ContainsKey(range.Text))
                                {
                                    Excel.Range flRange = ws.Cells[range.Row, flColIndex] as Excel.Range;
                                    flRange.Value = GlobalMembers.InstanceGlobalMembers.DictionaryDefaultFieldLength[range.Text];
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to add a hyperlink for a cell
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="range"></param>
        public static void AddHyperlinkToACell(Excel._Worksheet ws, Excel.Range range, RibbonMenu menu)
        {
            try
            {
                if (!string.IsNullOrEmpty(range.Text))
                {
                    if (menu.Label.ToString().Equals("Fields And Dropdowns"))
                    {
                        int ftColIndex = ExcelHelper.ConvertColumnNameToColumnIndex("FieldType", ws);
                        Excel.Range ftColRange = ws.Cells[range.Row, ftColIndex] as Excel.Range;
                        if ("Dropdown".Equals(ftColRange.Text))
                        {
                            int ddColIndex = ExcelHelper.ConvertColumnNameToColumnIndex("DropDownValues", ws);
                            Excel.Range ddRange = ws.Cells[range.Row, ddColIndex] as Excel.Range;
                            ws.Hyperlinks.Add(ddRange, "#'" + GlobalMembers.InstanceGlobalMembers.DropDownSheetName + "'!A3", "", "", "Drop Down Values");
                        }
                    }
                }
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Clear all the values in the log column for all the rows
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="oXL"></param>
        public static void ClearLogColumnsForAllTheRows(Excel._Worksheet sh)
        {
            for (int rowCount = 4; rowCount <= sh.UsedRange.Rows.Count; rowCount++)
            {
                Excel.Range r = sh.Cells[rowCount, 1] as Excel.Range;
                if (!string.IsNullOrEmpty(r.Text))
                {
                    r.ClearContents();
                    r.ClearComments();
                    r.ClearFormats();
                }
                else
                { break; }
            }

        }
        #endregion
        #region Reference Methods
        /// <summary>
        /// Adding a new sheet for drop downs
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="rTarget"></param>
        /// <param name="oWB"></param>
        public static void AddNewSheetForDropDown(Excel._Worksheet ws, Excel.Range rTarget, Excel.Workbook oWB)
        {
            int colIndex = 0, colCount = 1;
            try
            {
                colIndex = ConvertColumnNameToColumnIndex("FieldName", ws);
                Excel.Range r = ws.Cells[rTarget.Row, colIndex] as Excel.Range;
                Excel.Worksheet newWorksheet = oWB.Worksheets.Add(After: oWB.Sheets[oWB.Sheets.Count]);
                if (!CheckIfSheetExists(oWB, r.Text))
                {
                    newWorksheet.Name = r.Text;
                    Excel.Worksheet sheet = (Excel.Worksheet)oWB.Sheets[GlobalMembers.InstanceGlobalMembers.MetaDataSheetName];
                    sheet.Select(Type.Missing);
                    List<string> dropdownHeaders = GlobalMembers.InstanceGlobalMembers.ListDropdownHeaders;
                    foreach (var dropdownHeader in dropdownHeaders)
                    {
                        newWorksheet.Cells[1, colCount] = dropdownHeader;
                        r = newWorksheet.Cells[1, colCount] as Excel.Range;
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                        r.Font.Bold = true;
                        colCount++;
                    }
                    r.EntireColumn.AutoFit();
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        /// <summary>
        /// Check if sheet is existing or not
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private static bool CheckIfSheetExists(Excel.Workbook workbook, string sheetName)
        {
            bool isExists = false;
            Excel.Sheets sheets = workbook.Worksheets;
            foreach (Excel.Worksheet ws in sheets)
            {
                if (ws.Name.Equals(sheetName))
                { isExists = true; break; }
            }
            return isExists;
        }
        /// <summary>
        /// Add new sheet for grid field
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="rTarget"></param>
        /// <param name="oWB"></param>
        /// <param name="dictModulesList"></param>
        public static void AddNewSheetForGrid(Excel._Worksheet ws, Excel.Range rTarget, Excel.Workbook oWB, Dictionary<string, string> dictModulesList)
        {
            int colIndex = 0;
            try
            {
                Excel.Application application = (Excel.Application)oWB.Parent;
                ToggleExcelEvents(application, false);
                colIndex = ConvertColumnNameToColumnIndex("FieldName", ws);
                Excel.Range r = ws.Cells[rTarget.Row, colIndex] as Excel.Range;
                Excel.Worksheet newWorksheet = oWB.Worksheets.Add(After: oWB.Sheets[oWB.Sheets.Count]);
                newWorksheet.Name = r.Text;

                //DisplayMetaDataInformationOnSheet(newWorksheet, dictModulesList);

                Excel.Worksheet sheet = (Excel.Worksheet)oWB.Sheets[GlobalMembers.InstanceGlobalMembers.MetaDataSheetName];
                sheet.Select(Type.Missing);
                ToggleExcelEvents(application, true);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        /// <summary>
        /// Delete work sheets from workbook
        /// </summary>
        /// <param name="oWB"></param>
        public static void DeleteWorksheetFromWorkbook(Excel.Workbook oWB)
        {
            Excel.Sheets oSheets = oWB.Worksheets;
            foreach (Excel._Worksheet sh in oSheets)
            {
                //if (sh.Name != "Introduction" && sh.Name != "Configuration" && sh.Name != "MetaData" && sh.Name != "DropDownValues" && sh.Name != "Translation")
                if (!(sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.IntroductionSheetName))
                && (sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.ConfigurationSheetName))
                && (sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.MetaDataSheetName))
                && (sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.DropDownSheetName))
                && (sh.Name.Equals(GlobalMembers.InstanceGlobalMembers.TranslationSheetName))
                )
                { sh.Delete(); }
            }
        }
        /// <summary>
        /// Method used to display metadata information on sheet
        /// </summary>
        /// <param name="sh"></param>
        public static void DisplayMetaDataInformationOnSheet_Old1(Excel._Worksheet sh, Dictionary<string, string> dictModulesList)
        {
            int rowCount = 3, colCount = 2, startColumnIndex = 0;
            string mandatory = string.Empty;
            Excel.Range oRange;
            try
            {
                sh.Cells[rowCount - 1, 1] = "Log";
                sh.Cells[rowCount, 1] = "Log";
                ((Excel.Range)sh.Cells[rowCount - 2, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                ((Excel.Range)sh.Cells[rowCount, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                var distinctItems = GlobalMembers.InstanceGlobalMembers.ListMetadataData.GroupBy(x => x.TabName).Select(y => y.FirstOrDefault());
                foreach (var item in distinctItems)
                {
                    startColumnIndex = colCount;
                    var metaDataItems = GlobalMembers.InstanceGlobalMembers.ListMetadataData.Where(x => x.TabName == item.TabName.ToString());
                    foreach (var mdItem in metaDataItems)
                    {
                        oRange = sh.Cells[rowCount, colCount] as Excel.Range;
                        oRange.ClearComments();
                        sh.Cells[rowCount - 1, colCount] = mdItem.FieldName;

                        if (mdItem.Mandatory == "1")
                        {
                            mandatory = "*";
                            sh.Cells[rowCount, colCount] = mdItem.Description + " " + mandatory;
                            string v = oRange.Text;
                            oRange.Characters[v.LastIndexOf(' ') + 1, v.Length].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            sh.Cells[rowCount, colCount] = mdItem.Description;
                        }

                        if (!string.IsNullOrEmpty(mdItem.HelpText))
                        { oRange.AddComment(mdItem.HelpText); }

                        if (mdItem.IsDropDown == true)
                        {
                            if (mdItem.DropDownType.Equals("TrueFalse"))
                            { CreateDropDown(GlobalMembers.InstanceGlobalMembers.DictionaryTrueFalse, (Excel.Range)sh.Cells[rowCount + 1, colCount]); }
                            else if (mdItem.DropDownType.Equals("YesNo"))
                            { CreateDropDown(GlobalMembers.InstanceGlobalMembers.DictionaryYesNo, (Excel.Range)sh.Cells[rowCount + 1, colCount]); }
                            else if (mdItem.DropDownType.Equals("Other"))
                            {
                                if (mdItem.FieldName.Equals("FieldType"))
                                { CreateDropDown(GlobalMembers.InstanceGlobalMembers.DictionaryFieldTypes, (Excel.Range)sh.Cells[rowCount + 1, colCount]); }
                                else if (mdItem.FieldName.Equals("DataType"))
                                { CreateDropDown(GlobalMembers.InstanceGlobalMembers.DictionaryDataTypes, (Excel.Range)sh.Cells[rowCount + 1, colCount]); }
                                else if (mdItem.FieldName.Equals("refObjId"))
                                { CreateDropDown(dictModulesList, (Excel.Range)sh.Cells[rowCount + 1, colCount]); }
                                else if (mdItem.Equals("AttachmentFileType"))
                                { CreateDropDown(GlobalMembers.InstanceGlobalMembers.DictionaryAttachmentFileTypes, (Excel.Range)sh.Cells[rowCount + 1, colCount]); }
                            }
                        }
                        colCount++;
                    }
                    sh.Cells[rowCount - 2, colCount - 1] = item.TabName.ToString();
                    var mergeRange = sh.Range[sh.Cells[1, startColumnIndex], sh.Cells[1, colCount - 1]];
                    mergeRange.Merge(true);
                    mergeRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(item.Colour);
                    mergeRange.Cells.Font.Bold = true;
                    mergeRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    mergeRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                oRange = sh.Cells[2, colCount] as Excel.Range;
                oRange.EntireRow.Hidden = true;

                oRange = sh.Range[sh.Cells[1, 1], sh.Cells[3, colCount - 1]] as Excel.Range;
                oRange.EntireColumn.AutoFit();

                oRange = sh.Range[sh.Cells[1, 1], sh.Cells[1, colCount - 1]] as Excel.Range;
                oRange.RowHeight = 20;

                oRange = sh.Range[sh.Cells[3, 1], sh.Cells[3, colCount - 1]] as Excel.Range;
                oRange.RowHeight = 20;
                oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                oRange.Cells.Font.Bold = true;
                oRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                var cRange = sh.Range[sh.Cells[4, 1], sh.Cells[4, colCount]] as Excel.Range;
                var dRange = sh.Range[sh.Cells[5, 1], sh.Cells[1000, colCount]] as Excel.Range;
                cRange.Select();
                cRange.Copy(dRange);
                var cell1 = sh.Cells[1, 1] as Excel.Range;
                cell1.Select();
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        /// <summary>
        /// Method used to create table for offline save
        /// </summary>
        public static void CreateTableForOfflineSave()
        {
            string[] columnName = new string[] { "FieldName", "FieldDescription", "FieldType", "DataType", "FieldLength", "helpText", "outputLength", "mandatory", "parentField", "picklistDependency", "nounFld", "refObjId", "refFieldId", "searchField", "isRefField", "descFld", "Keys", "permissionFld", "numberSettingCrit", "workflowFld", "workflowCriteria", "tabCriteriaFld", "srchEngine", "isTransientField", "suggestion", "isCompBased", "isCompleteness", "checkList", "DecimalPlace", "tnccheck", "AttachmentSize", "AttachmentFileType", "optionLimit", "textAreaLength", "textAreaWidth", "defaultDate", "futureDate", "pastDate", "gridDisplay", "defDisplay" };
            string[] dataType = new string[] { "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT", "TEXT" };
            GlobalMembers.InstanceGlobalMembers.SqliteDatabase.CreateTable("SavedExcelData", columnName, dataType);
        }
        #endregion
    }
}