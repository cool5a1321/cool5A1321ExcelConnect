using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Prospecta.ConnektHub.Helpers;
using Prospecta.ConnektHub.Models;
using Prospecta.ConnektHub.Services.HttpService;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace Prospecta.ConnektHub.Controllers
{
    public class RibbonController
    {
        #region Global Variables for the class
        private Dictionary<string, string> dictModulesList = new Dictionary<string, string>();
        IHttpRequest _httpRequest;
        #endregion
        #region Constructors
        /// <summary>
        /// Parameterized constructor
        /// </summary>
        /// <param name="httpRequest"></param>
        public RibbonController(IHttpRequest httpRequest)
        { _httpRequest = httpRequest; }
        #endregion
        #region Private Methods
        /// <summary>
        /// Method used to convert the sheet to data table
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="startRow"></param>
        /// <returns></returns>
        public System.Data.DataTable ConvertSheetDataToDataTable(_Worksheet sh, int startRow)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            int colCount = sh.UsedRange.Columns.Count;
            int rowCount = sh.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            for (int cCnt = 1; cCnt <= colCount; cCnt++)
            {
                Range r = sh.Cells[1, cCnt] as Range;
                dt.Columns.Add(r.Text, typeof(System.String));
            }

            for (int rCnt = startRow; rCnt <= rowCount; rCnt++)
            {
                DataRow dr = dt.NewRow();
                for (int cCnt = 1; cCnt <= colCount; cCnt++)
                {
                    Range columnName = sh.Cells[1, cCnt] as Range;
                    Range columnValue = sh.Cells[rCnt, cCnt] as Range;
                    dr[columnName.Text] = columnValue.Text;
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        #endregion
        #region Public Methods
        /// <summary>
        /// Method used to build the JSON for translated descriptions
        /// </summary>
        /// <param name="sh"></param>
        public void BuildTranlatedDataJsonAndExport(_Worksheet sh, Dictionary<string, string> dictLanguage)
        {
            try
            {
                Dictionary<string, object> dictParent = new Dictionary<string, object>();
                int rowCount = 3;
                System.Data.DataTable dtTranslatedData = ConvertSheetDataToDataTable(sh, rowCount);
                int totalRows = sh.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                for (int rCnt = rowCount; rCnt <= totalRows; rCnt++)
                {
                    Range fieldName = sh.Cells[rCnt, 1] as Range;
                    if (!dictParent.ContainsKey(fieldName.Text))
                    {
                        string filterCondition = string.Format("FieldId = '{0}'", new string[] { fieldName.Text });
                        DataRow[] drArray = dtTranslatedData.Select(filterCondition);
                        List<Dictionary<string, string>> lstDictionary = new List<Dictionary<string, string>>();
                        foreach (DataRow dr in drArray)
                        {
                            Dictionary<string, string> dictTranslatedValues = new Dictionary<string, string>();
                            foreach (DataColumn dc in dr.Table.Columns)
                            {
                                if (dc.ColumnName.ToString().Equals("Language"))
                                {
                                    string languageName = dr[dc].ToString();
                                    string languageCode = dictLanguage[languageName];
                                    dictTranslatedValues.Add(dc.ColumnName.ToString(), languageCode);
                                }
                                else
                                { dictTranslatedValues.Add(dc.ColumnName.ToString(), dr[dc].ToString()); }
                            }
                            lstDictionary.Add(dictTranslatedValues);
                        }
                        dictParent.Add(fieldName.Text, lstDictionary);
                    }
                }
                string jsonString = JsonConvert.SerializeObject(dictParent);
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to build the JSON for dropdown sheet
        /// </summary>
        /// <param name="sh"></param>
        public void BuildDropDownJsonAndExport(_Worksheet sh, Dictionary<string, string> dictLanguage)
        {
            try
            {
                Dictionary<string, object> dictParent = new Dictionary<string, object>();
                int totalRows = sh.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                int rowCount = 3;
                System.Data.DataTable dtDropDownData = ConvertSheetDataToDataTable(sh, rowCount);
                for (int rCnt = rowCount; rCnt <= totalRows; rCnt++)
                {
                    Range fieldName = sh.Cells[rCnt, 1] as Range;
                    if (!dictParent.ContainsKey(fieldName.Text))
                    {
                        string filterCondition = string.Format("FieldId = '{0}'", new string[] { fieldName.Text });
                        DataRow[] drArray = dtDropDownData.Select(filterCondition);
                        List<Dictionary<string, string>> lstDictionary = new List<Dictionary<string, string>>();
                        foreach (DataRow dr in drArray)
                        {
                            Dictionary<string, string> dictDropDownValues = new Dictionary<string, string>();
                            foreach (DataColumn dc in dr.Table.Columns)
                            {
                                if (dc.ColumnName.ToString().Equals("Language"))
                                {
                                    string languageName = dr[dc].ToString();
                                    string languageCode = dictLanguage[languageName];
                                    dictDropDownValues.Add(dc.ColumnName.ToString(), languageCode);
                                }
                                else
                                { dictDropDownValues.Add(dc.ColumnName.ToString(), dr[dc].ToString()); }
                            }
                            lstDictionary.Add(dictDropDownValues);
                        }
                        dictParent.Add(fieldName.Text, lstDictionary);
                    }
                }
                string jsonString = JsonConvert.SerializeObject(dictParent);
            }
            catch (Exception ex)
            { throw (ex); }
        }
        /// <summary>
        /// Method used to validate the metadata sheet
        /// </summary>
        /// <param name="rowCount"></param>
        /// <param name="colCount"></param>
        /// <param name="oSheet"></param>
        public void ValidateExcelData(int rowCount, int colCount, _Worksheet oSheet)
        {
            string errorMessage = string.Empty, filterCondition = string.Empty;
            int rowCnt = 0, result = 0;
            rowCnt = rowCount;
            for (int colCnt = 2; colCnt <= colCount; colCnt++)
            {
                Range headerRange = oSheet.Cells[2, colCnt] as Range;
                Range dataRange = oSheet.Cells[rowCount, colCnt] as Range;
                string dataValue = dataRange.Text;
                filterCondition = string.Format("FieldName = '{0}'", new string[] { headerRange.Text });
                var dataRowArray = GlobalMembers.InstanceGlobalMembers.DtMetadataData.Select(filterCondition);
                for (int iCnt = 0; iCnt < dataRowArray.Length; iCnt++)
                {
                    if (string.IsNullOrEmpty(dataValue) && dataRowArray[iCnt]["Mandatory"].ToString().Equals("1"))
                    {
                        if (string.IsNullOrEmpty(errorMessage))
                        { errorMessage = "Please enter value for mandatory field " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text; }
                        else
                        { errorMessage += ", Please enter value for mandatory field " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text; }
                        dataRange.Borders.Color = 0x0000FF;
                    }
                    else
                    {
                        if (dataRowArray[iCnt]["FieldName"].ToString().Equals("FieldName"))
                        {
                            string fieldName = dataRange.Text;
                            if (GlobalMembers.InstanceGlobalMembers.ListSqlKeyWords.Contains(fieldName.ToUpper()))
                            {
                                if (string.IsNullOrEmpty(errorMessage))
                                    errorMessage = "Field Name cannot be any sql keyword";
                                else
                                    errorMessage += ", Field Name cannot be any sql keyword";
                            }

                            if (Regex.IsMatch(fieldName, @"\s"))
                            {
                                if (string.IsNullOrEmpty(errorMessage))
                                    errorMessage = "Field Name cannot have spaces";
                                else
                                    errorMessage += ", Field Name cannot have spaces";
                            }

                            if (Regex.IsMatch(fieldName, @"^\d+"))
                            {
                                if (string.IsNullOrEmpty(errorMessage))
                                    errorMessage = "Field Name should not start with numbers";
                                else
                                    errorMessage += ", Field Name should not start with numbers";
                            }

                            if (rowCnt == 4)
                            {
                                ExcelHelper.DuplicacyCheckInFieldName(oSheet);
                            }
                        }
                        if (dataRowArray[iCnt]["FieldName"].ToString().Equals("refObjId"))
                        {
                            if (!string.IsNullOrEmpty(dataValue))
                            {
                                int colIndex = ExcelHelper.ConvertColumnNameToColumnIndex("refFieldId", oSheet);
                                var r = oSheet.Cells[rowCnt, colIndex] as Range;
                                if (r.Text == "")
                                {
                                    if (string.IsNullOrEmpty(errorMessage))
                                        errorMessage = "Reference field id is mandatory";
                                    else
                                        errorMessage += ", Reference field id is mandatory";
                                }
                                r.Borders.Color = 0x0000FF;
                            }
                        }

                        if (dataRowArray[iCnt]["FieldName"].ToString().Equals("FieldDescription"))
                        {
                            if (!string.IsNullOrEmpty(dataValue))
                            {
                                if (dataValue.Length > 100)
                                {
                                    if (string.IsNullOrEmpty(errorMessage))
                                        errorMessage = "Length of " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " column should be less than 100.";
                                    else
                                        errorMessage += ", Length of " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " column should be less than 100.";
                                    dataRange.Borders.Color = 0x0000FF;
                                }
                            }
                        }

                        if (dataRowArray[iCnt]["FieldName"].ToString().Equals("FieldLength"))
                        {
                            if (!string.IsNullOrEmpty(dataValue))
                            {
                                int.TryParse(dataValue, out result);
                                if (result.Equals(0))
                                {
                                    if (string.IsNullOrEmpty(errorMessage))
                                        errorMessage = "Please enter numeric value in the " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " field.";
                                    else
                                        errorMessage += ", Please enter numeric value in the " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " field.";
                                    dataRange.Borders.Color = 0x0000FF;
                                }
                                else
                                {
                                    int fieldLength = Convert.ToInt32(dataRange.Text);
                                    if (fieldLength > 4000)
                                    {
                                        if (string.IsNullOrEmpty(errorMessage))
                                            errorMessage = "Field Length should not be greater than 4000.";
                                        else
                                            errorMessage += ", Field Length should not be greater than 4000.";
                                        dataRange.Borders.Color = 0x0000FF;
                                    }
                                }
                            }
                        }

                        if (dataRowArray[iCnt]["FieldName"].ToString().Equals("FieldType"))
                        {
                            if (
                                dataRange.Text == "Radio" ||
                                dataRange.Text == "Module Reference Id" ||
                                dataRange.Text == "TextArea"
                                )
                            {
                                if (GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies.ContainsKey(dataRange.Text))
                                {
                                    string[] arrayDependencies = GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies[dataRange.Text].Split(',');
                                    for (int i = 0; i < arrayDependencies.Length; i++)
                                    {
                                        string val = arrayDependencies[i].Trim();
                                        int colIndex = ExcelHelper.ConvertColumnNameToColumnIndex(val, oSheet);
                                        var r = oSheet.Cells[2, colIndex] as Range;

                                        if (r.Text == "optionLimit")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Option limit is mandatory";
                                                else
                                                    errorMessage += ", Option limit is mandatory";
                                            }
                                            else
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Option limit should be numeric only";
                                                    else
                                                        errorMessage += ", Option limit should be numeric only";
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(errorMessage))
                                            { r.Borders.Color = 0x0000FF; }
                                        }

                                        if (r.Text == "refObjId")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Reference Object Id is mandatory if " + dataRange.Text + " is selected";
                                                else
                                                    errorMessage += ", Reference Object Id is mandatory if " + dataRange.Text + " is selected";
                                                r.Borders.Color = 0x0000FF;
                                            }
                                        }

                                        if (r.Text == "textAreaLength")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Text Area Length is mandatory";
                                                else
                                                    errorMessage += ", Text Area Length limit is mandatory";
                                            }
                                            else
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Text Area Length should be numeric only";
                                                    else
                                                        errorMessage += ", Text Area Length should be numeric only";
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(errorMessage))
                                            { r.Borders.Color = 0x0000FF; }
                                        }
                                        if (r.Text == "textAreaWidth")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Text Area Width is mandatory";
                                                else
                                                    errorMessage += ", Text Area Width limit is mandatory";
                                            }
                                            else
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Text Area Width should be numeric only";
                                                    else
                                                        errorMessage += ", Text Area Width should be numeric only";
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(errorMessage))
                                            { r.Borders.Color = 0x0000FF; }
                                        }

                                        if (r.Text == "AttachmentFileType")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Attachment File Type should not be blank";
                                                else
                                                    errorMessage += ", Attachment File Type should not be blank";
                                                r.Borders.Color = 0x0000FF;
                                            }
                                        }

                                        if (r.Text == "AttachmentSize")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text != "")
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Attachment Size should be numeric only";
                                                    else
                                                        errorMessage += ", Attachment Size should be numeric only";
                                                    r.Borders.Color = 0x0000FF;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            var cells = oSheet.Cells[rowCnt, 1] as Range;

            if (string.IsNullOrEmpty(cells.Text))
            {
                cells.ClearComments();
                cells.Value = "";
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    cells.Value = "IN VALID";
                    cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                    cells.AddComment(errorMessage);
                }
                else
                {
                    cells.Value = "VALID";
                    cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                    cells.ClearComments();
                }
                cells.EntireColumn.AutoFit();
            }
        }
        /// <summary>
        /// Method used to Build and export field list json to server
        /// </summary>
        /// <param name="rowCount"></param>
        /// <param name="colCount"></param>
        /// <param name="oSheet"></param>
        /// <param name="objectType"></param>
        /// <param name="userId"></param>
        public void BuildAndExportFieldsJson(int rowCount, int colCount, _Worksheet oSheet, string objectType, string userId)
        {
            int rowCnt = rowCount;

            string dropValue = string.Empty;
            Range oRange = oSheet.Cells[rowCnt, 1] as Range;
            string logValue = Convert.ToString(oRange.Text);
            if (logValue.ToLower() == "valid")
            {
                oRange.Value2 = "";
                oRange.ClearComments();
                oRange.ClearFormats();

                Dictionary<string, object> dictParent = new Dictionary<string, object>();
                Dictionary<string, string> dict = new Dictionary<string, string>();
                for (int colCnt = 2; colCnt <= colCount; colCnt++)
                {
                    Range headerRange = oSheet.Cells[2, colCnt] as Range;
                    Range dataRange = oSheet.Cells[rowCnt, colCnt] as Range;
                    string filterCondition = string.Empty, dataValue = string.Empty;
                    filterCondition = string.Format("FieldName = '{0}'", new string[] { headerRange.Text });
                    var dataRowArray = GlobalMembers.InstanceGlobalMembers.DtMetadataData.Select(filterCondition);
                    for (int i = 0; i < dataRowArray.Length; i++)
                    {
                        if (dataRowArray[i]["IsDropDown"].ToString().Equals("1"))
                        {
                            if (dataRowArray[i]["DropDownType"].ToString().Equals("Other"))
                            {
                                if (!string.IsNullOrEmpty(dataRange.Text))
                                {
                                    if (dataRowArray[i]["FieldName"].Equals("FieldType"))
                                    {
                                        dict.Add(dataRowArray[i]["FieldName"].ToString(), ExcelHelper.GetValueFromCodeForDropDown(GlobalMembers.InstanceGlobalMembers.DtFieldTypeValues, dataRange.Text));
                                    }
                                    else if (dataRowArray[i]["FieldName"].Equals("DataType"))
                                    {
                                        dict.Add(dataRowArray[i]["FieldName"].ToString(), ExcelHelper.GetValueFromCodeForDropDown(GlobalMembers.InstanceGlobalMembers.DtDataTypeValues, dataRange.Text));
                                    }
                                    else if (dataRowArray[i]["FieldName"].Equals("refObjId"))
                                    {
                                        dict.Add(dataRowArray[i]["FieldName"].ToString(), dictModulesList[dataRange.Text]);
                                    }
                                    else if (dataRowArray[i]["FieldName"].Equals("AttachmentFileType"))
                                    {
                                        dict.Add(dataRowArray[i]["FieldName"].ToString(), ExcelHelper.GetValueFromCodeForDropDown(GlobalMembers.InstanceGlobalMembers.DtAttachmentTypeValues, dataRange.Text));
                                    }
                                }
                                else
                                { dict.Add(dataRowArray[i]["FieldName"].ToString(), dataRange.Text); }
                            }
                            else if (dataRowArray[i]["DropDownType"].ToString().Equals("YesNo"))
                            {
                                if (!string.IsNullOrEmpty(dataRange.Text))
                                { dataValue = dataRange.Text; }
                                else
                                {
                                    if (dataRowArray[i]["FieldName"].ToString().Equals("nounFld") || dataRowArray[i]["FieldName"].ToString().Equals("picklistDependency"))
                                    { dataValue = string.Empty; }
                                    else if (dataRowArray[i]["FieldName"].ToString().Equals("defDisplay"))
                                    { dataValue = "Yes"; }
                                    else
                                    { dataValue = "No"; }
                                }

                                if (!dataValue.Equals(string.Empty))
                                { dict.Add(dataRowArray[i]["FieldName"].ToString(), ExcelHelper.GetValueFromCodeForDropDown(GlobalMembers.InstanceGlobalMembers.DtYesNoValues, dataValue)); }
                                else
                                { dict.Add(dataRowArray[i]["FieldName"].ToString(), string.Empty); }
                            }
                            else if (dataRowArray[i]["DropDownType"].ToString().Equals("TrueFalse"))
                            {
                                if (!string.IsNullOrEmpty(dataRange.Text))
                                { dataValue = dataRange.Text; }
                                else
                                { dataValue = "false"; }
                                dict.Add(dataRowArray[i]["FieldName"].ToString(), ExcelHelper.GetValueFromCodeForDropDown(GlobalMembers.InstanceGlobalMembers.DtTrueFalseValues, dataValue));
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(dataRange.Text))
                            { dataValue = dataRange.Text; }
                            else
                            {
                                if (dataRowArray[i]["FieldName"].ToString().Equals("optionLimit"))
                                { dataValue = "1"; }
                                else if (dataRowArray[i]["FieldName"].ToString().Equals("outputLength"))
                                { dataValue = "200"; }
                            }
                            dict.Add(dataRowArray[i]["FieldName"].ToString(), dataValue);
                        }
                    }
                }
                dict.Add("StructureId", "");
                dict.Add("objectId", objectType);
                dict.Add("loc_Type", "");
                dict.Add("ObjectMap", "");
                dict.Add("formula", "");
                dict.Add("seperator", "");
                dict.Add("AttachmentWidth", "");
                dict.Add("AttachmentHeight", "");
                dict.Add("DecimalField", "");
                dict.Add("noOfFields", "");
                dict.Add("imagemarker", "1");
                dict.Add("groupDetails", "{}");
                dict.Add("financialYrStrt", "0");
                dict.Add("financialYrEnd", "0");
                dict.Add("financialYrVal", "0");
                dict.Add("apiField", "");
                dict.Add("eventId", "create");
                dict.Add("apiSno", "");
                dict.Add("colorType", "");
                dict.Add("matrixFldX", "");
                dict.Add("matrixFldY", "");
                dict.Add("aggFieldCriteriaMap", "{}");
                dictParent.Add("fieldData", dict);

                string url = "restObjectList/createField/" + objectType + "/" + userId + "/" + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowCnt, 2]).Text;
                var jsonString = JsonConvert.SerializeObject(dictParent);
                string responseJson = _httpRequest.HttpPost(url, jsonString, false);
                var response = JsonConvert.DeserializeObject<WebServiceResponse>(responseJson);
                string serviceStatus = response.RESPONSE_STATUS;

                if ("SUCCESS".Equals(serviceStatus.ToUpper()))
                {
                    oRange.Value = serviceStatus.ToUpper();
                    oRange.AddComment("Field Created Successfully");
                    oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                }
                else
                {
                    oRange.Value = serviceStatus.ToUpper();
                    oRange.AddComment(response.ERROR_MSG);
                    oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                }
            }
        }
        #endregion
        #region code for reference
        /// <summary>
        /// Method used to validate excel data
        /// </summary>
        /// <param name="rowCount"></param>
        /// <param name="colCount"></param>
        /// <param name="oSheet"></param>
        public void ValidateExcelData_Old(int rowCount, int colCount, _Worksheet oSheet)
        {
            int result = 0;
            int rowCnt = rowCount;

            string errorMessage = string.Empty;
            for (int colCnt = 1; colCnt <= colCount; colCnt++)
            {
                Range headerRange = oSheet.Cells[2, colCnt] as Range;
                Range dataRange = oSheet.Cells[rowCnt, colCnt] as Range;

                var metaDataItem = GlobalMembers.InstanceGlobalMembers.ListMetadataData.Where(x => x.FieldName == headerRange.Text);
                foreach (var item in metaDataItem)
                {
                    if (dataRange.Text == "" && item.Mandatory == "1")
                    {
                        if (string.IsNullOrEmpty(errorMessage))
                        { errorMessage = "Please enter value for mandatory field " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text; }
                        else
                        { errorMessage += ", Please enter value for mandatory field " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text; }
                        dataRange.Borders.Color = 0x0000FF;
                    }
                    else
                    {

                        if (item.FieldName == "FieldName")
                        {
                            string fieldName = dataRange.Text;
                            if (GlobalMembers.InstanceGlobalMembers.ListSqlKeyWords.Contains(fieldName.ToUpper()))
                            {
                                if (string.IsNullOrEmpty(errorMessage))
                                    errorMessage = "Field Name cannot be any sql keyword";
                                else
                                    errorMessage += ", Field Name cannot be any sql keyword";
                            }

                            if (Regex.IsMatch(fieldName, @"\s"))
                            {
                                if (string.IsNullOrEmpty(errorMessage))
                                    errorMessage = "Field Name cannot have spaces";
                                else
                                    errorMessage += ", Field Name cannot have spaces";
                            }

                            if (Regex.IsMatch(fieldName, @"^\d+"))
                            {
                                if (string.IsNullOrEmpty(errorMessage))
                                    errorMessage = "Field Name should not start with numbers";
                                else
                                    errorMessage += ", Field Name should not start with numbers";
                            }

                            if (rowCnt == 4)
                            {
                                ExcelHelper.DuplicacyCheckInFieldName(oSheet);
                            }
                        }

                        if (item.FieldName == "refObjId")
                        {
                            if (dataRange.Text != "")
                            {
                                int colIndex = ExcelHelper.ConvertColumnNameToColumnIndex("refFieldId", oSheet);
                                var r = oSheet.Cells[rowCnt, colIndex] as Range;
                                if (r.Text == "")
                                {
                                    if (string.IsNullOrEmpty(errorMessage))
                                        errorMessage = "Reference field id is mandatory";
                                    else
                                        errorMessage += ", Reference field id is mandatory";
                                }
                                r.Borders.Color = 0x0000FF;
                            }
                        }

                        if (item.FieldName == "FieldDescription")
                        {
                            if (dataRange.Text != "")
                            {
                                if (((string)dataRange.Text).Length > 100)
                                {
                                    if (string.IsNullOrEmpty(errorMessage))
                                        errorMessage = "Length of " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " column should be less than 100.";
                                    else
                                        errorMessage += ", Length of " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " column should be less than 100.";
                                    dataRange.Borders.Color = 0x0000FF;
                                }
                            }
                        }

                        if (item.FieldName == "FieldLength")
                        {
                            if (!string.IsNullOrEmpty(dataRange.Text))
                            {
                                int.TryParse(dataRange.Text, out result);
                                if (result.Equals(0))
                                {
                                    if (string.IsNullOrEmpty(errorMessage))
                                        errorMessage = "Please enter numeric value in the " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " field.";
                                    else
                                        errorMessage += ", Please enter numeric value in the " + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[2, colCnt]).Text + " field.";
                                    dataRange.Borders.Color = 0x0000FF;
                                }
                                else
                                {
                                    int fieldLength = Convert.ToInt32(dataRange.Text);
                                    if (fieldLength > 4000)
                                    {
                                        if (string.IsNullOrEmpty(errorMessage))
                                            errorMessage = "Field Length should not be greater than 4000.";
                                        else
                                            errorMessage += ", Field Length should not be greater than 4000.";
                                        dataRange.Borders.Color = 0x0000FF;
                                    }
                                }
                            }
                        }

                        if (item.FieldName == "FieldType")
                        {
                            if (
                                dataRange.Text == "Radio" ||
                                dataRange.Text == "Module Reference Id" ||
                                dataRange.Text == "TextArea"
                                )
                            {
                                if (GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies.ContainsKey(dataRange.Text))
                                {
                                    string[] arrayDependencies = GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies[dataRange.Text].Split(',');
                                    for (int i = 0; i < arrayDependencies.Length; i++)
                                    {
                                        string val = arrayDependencies[i].Trim();
                                        int colIndex = ExcelHelper.ConvertColumnNameToColumnIndex(val, oSheet);
                                        var r = oSheet.Cells[2, colIndex] as Range;

                                        if (r.Text == "optionLimit")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Option limit is mandatory";
                                                else
                                                    errorMessage += ", Option limit is mandatory";
                                            }
                                            else
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Option limit should be numeric only";
                                                    else
                                                        errorMessage += ", Option limit should be numeric only";
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(errorMessage))
                                            { r.Borders.Color = 0x0000FF; }
                                        }

                                        if (r.Text == "refObjId")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Reference Object Id is mandatory if " + dataRange.Text + " is selected";
                                                else
                                                    errorMessage += ", Reference Object Id is mandatory if " + dataRange.Text + " is selected";
                                                r.Borders.Color = 0x0000FF;
                                            }
                                        }

                                        if (r.Text == "textAreaLength")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Text Area Length is mandatory";
                                                else
                                                    errorMessage += ", Text Area Length limit is mandatory";
                                            }
                                            else
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Text Area Length should be numeric only";
                                                    else
                                                        errorMessage += ", Text Area Length should be numeric only";
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(errorMessage))
                                            { r.Borders.Color = 0x0000FF; }
                                        }
                                        if (r.Text == "textAreaWidth")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Text Area Width is mandatory";
                                                else
                                                    errorMessage += ", Text Area Width limit is mandatory";
                                            }
                                            else
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Text Area Width should be numeric only";
                                                    else
                                                        errorMessage += ", Text Area Width should be numeric only";
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(errorMessage))
                                            { r.Borders.Color = 0x0000FF; }
                                        }

                                        if (r.Text == "AttachmentFileType")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text == "")
                                            {
                                                if (string.IsNullOrEmpty(errorMessage))
                                                    errorMessage = "Attachment File Type should not be blank";
                                                else
                                                    errorMessage += ", Attachment File Type should not be blank";
                                                r.Borders.Color = 0x0000FF;
                                            }
                                        }

                                        if (r.Text == "AttachmentSize")
                                        {
                                            r = oSheet.Cells[rowCnt, colIndex] as Range;
                                            if (r.Text != "")
                                            {
                                                int.TryParse(r.Text, out result);
                                                if (result == 0)
                                                {
                                                    if (string.IsNullOrEmpty(errorMessage))
                                                        errorMessage = "Attachment Size should be numeric only";
                                                    else
                                                        errorMessage += ", Attachment Size should be numeric only";
                                                    r.Borders.Color = 0x0000FF;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            var cells = oSheet.Cells[rowCnt, 1] as Range;

            if (string.IsNullOrEmpty(cells.Text))
            {
                cells.ClearComments();
                cells.Value = "";
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    cells.Value = "IN VALID";
                    cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                    cells.AddComment(errorMessage);
                }
                else
                {
                    cells.Value = "VALID";
                    cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                    cells.ClearComments();
                }
                cells.EntireColumn.AutoFit();
            }
        }
        /// <summary>
        /// Build and export json to server
        /// </summary>
        /// <param name="rowCount"></param>
        /// <param name="colCount"></param>
        /// <param name="oSheet"></param>
        /// <param name="objectType"></param>
        /// <param name="userId"></param>
        public void BuildAndExportJson_old(int rowCount, int colCount, _Worksheet oSheet, string objectType, string userId)
        {
            int rowCnt = rowCount;
            string dropValue = string.Empty;
            Range oRange = oSheet.Cells[rowCnt, 1] as Range;
            oRange.Value2 = "";
            oRange.ClearComments();
            oRange.ClearFormats();

            Dictionary<string, object> dictParent = new Dictionary<string, object>();
            Dictionary<string, string> dict = new Dictionary<string, string>();
            for (int colCnt = 2; colCnt <= colCount; colCnt++)
            {
                Range headerRange = oSheet.Cells[2, colCnt] as Range;
                Range dataRange = oSheet.Cells[rowCnt, colCnt] as Range;
                var metaDataItem = GlobalMembers.InstanceGlobalMembers.ListMetadataData.Where(x => x.FieldName == headerRange.Text);
                foreach (var item in metaDataItem)
                {
                    if (item.IsDropDown)
                    {
                        if (item.DropDownType == "Other")
                        {
                            if (!string.IsNullOrEmpty(dataRange.Text))
                            {
                                if (item.FieldName.Equals("FieldType"))
                                { dict.Add(item.FieldName, GlobalMembers.InstanceGlobalMembers.DictionaryFieldTypes[dataRange.Text]); }
                                else if (item.FieldName.Equals("DataType"))
                                { dict.Add(item.FieldName, GlobalMembers.InstanceGlobalMembers.DictionaryDataTypes[dataRange.Text]); }
                                else if (item.FieldName.Equals("refObjId"))
                                { dict.Add(item.FieldName, dictModulesList[dataRange.Text]); }
                                else if (item.FieldName.Equals("AttachmentFileType"))
                                { dict.Add(item.FieldName, GlobalMembers.InstanceGlobalMembers.DictionaryAttachmentFileTypes[dataRange.Text]); }
                            }
                            else
                            { dict.Add(item.FieldName, dataRange.Text); }
                        }
                        else if (item.DropDownType == "YesNo")
                        {
                            if (!string.IsNullOrEmpty(dataRange.Text))
                            { dict.Add(item.FieldName, GlobalMembers.InstanceGlobalMembers.DictionaryYesNo[dataRange.Text]); }
                            else
                            {
                                if (item.FieldName.Equals("nounFld") || item.FieldName.Equals("picklistDependency"))
                                { dict.Add(item.FieldName, string.Empty); }
                                else if (item.FieldName.Equals("defDisplay"))
                                { dict.Add(item.FieldName, "1"); }
                                else
                                { dict.Add(item.FieldName, "0"); }
                            }
                        }
                        else if (item.DropDownType == "TrueFalse")
                        {
                            if (!string.IsNullOrEmpty(dataRange.Text))
                            { dict.Add(item.FieldName, GlobalMembers.InstanceGlobalMembers.DictionaryTrueFalse[((string)dataRange.Text).ToLower()]); }
                            else
                            { dict.Add(item.FieldName, "false"); }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(dataRange.Text))
                        { dict.Add(item.FieldName, dataRange.Text); }
                        else
                        {
                            if (item.FieldName.Equals("optionLimit"))
                            { dict.Add(item.FieldName, "1"); }
                            else if (item.FieldName.Equals("outputLength"))
                            { dict.Add(item.FieldName, "200"); }
                        }
                    }
                }
            }
            dict.Add("StructureId", "");
            dict.Add("objectId", objectType);
            dict.Add("loc_Type", "");
            dict.Add("ObjectMap", "");
            dict.Add("formula", "");
            dict.Add("seperator", "");
            dict.Add("AttachmentWidth", "");
            dict.Add("AttachmentHeight", "");
            dict.Add("DecimalField", "");
            dict.Add("noOfFields", "");
            dict.Add("imagemarker", "1");
            dict.Add("groupDetails", "{}");
            dict.Add("financialYrStrt", "0");
            dict.Add("financialYrEnd", "0");
            dict.Add("financialYrVal", "0");
            dict.Add("apiField", "");
            dict.Add("eventId", "create");
            dict.Add("apiSno", "");
            dict.Add("colorType", "");
            dict.Add("matrixFldX", "");
            dict.Add("matrixFldY", "");
            dict.Add("aggFieldCriteriaMap", "{}");
            dictParent.Add("fieldData", dict);

            string url = "restObjectList/createField/" + objectType + "/" + userId + "/" + ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowCnt, 2]).Text;
            var jsonString = JsonConvert.SerializeObject(dictParent);
            string responseJson = _httpRequest.HttpPost(url, jsonString, false);
            var response = JsonConvert.DeserializeObject<WebServiceResponse>(responseJson);
            string serviceStatus = response.RESPONSE_STATUS;

            if ("SUCCESS".Equals(serviceStatus.ToUpper()))
            {
                oRange.Value = serviceStatus.ToUpper();
                oRange.AddComment("Field Created Successfully");
                oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            }
            else
            {
                oRange.Value = serviceStatus.ToUpper();
                oRange.AddComment(response.ERROR_MSG);
                oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
            }
        }
        #endregion
    }
}