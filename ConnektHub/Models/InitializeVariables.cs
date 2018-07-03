using ConnektHub.Models;
using Prospecta.ConnektHub.Helpers;
using System.Collections.Generic;
using System.Drawing;

namespace Prospecta.ConnektHub.Models
{
    public class InitializeVariables
    {
        #region Private Methods
        /// <summary>
        /// Field Type Dropdown
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeFieldTypes()
        {
            Dictionary<string, string> dictFieldTypes = new Dictionary<string, string>
            {
                { "Text", "0" },
                { "Dropdown", "1" },
                { "Checkbox", "2" },
                { "Radio", "4" },
                { "TextArea", "22" },
                { "Date", "DATS" },
                { "Time", "TIMS" },
                { "Date and Time", "DTMS" },
                { "Module Reference Id", "30" }
            };


            /* To be used later (Don't delete)
            dictFieldTypes.Add("Grid", "15");
            dictFieldTypes.Add("User Selection", "37");
            dictFieldTypes.Add("Location Reference", "29");
            dictFieldTypes.Add("Attachment", "28");
            dictFieldTypes.Add("Digital Sign", "44");
            dictFieldTypes.Add("Aggregation", "45");
            dictFieldTypes.Add("Translation Table", "39");
            dictFieldTypes.Add("HTML", "31");
            dictFieldTypes.Add("Calculation", "25");
            dictFieldTypes.Add("URL", "33");
            dictFieldTypes.Add("Position Type", "34");
            dictFieldTypes.Add("Tree", "5");
            dictFieldTypes.Add("Noun Type", "18");
            dictFieldTypes.Add("Work Flow", "10");
            dictFieldTypes.Add("Rejection Type", "38");
            dictFieldTypes.Add("Ativate/ Deactivate", "35");
            dictFieldTypes.Add("GeoLocation", "40");
            dictFieldTypes.Add("Group", "14");
            dictFieldTypes.Add("Email", "EMAIL");
            dictFieldTypes.Add("Password", "PASS");
            */

            return dictFieldTypes;
        }
        /// <summary>
        /// YesNo Dropdown
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeYesNo()
        {
            Dictionary<string, string> dictYesNo = new Dictionary<string, string>
            {
                { "No", "0" },
                { "Yes", "1" }
            };
            return dictYesNo;
        }
        /// <summary>
        ///  Attachment type drop down 
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeAttachmentTypes()
        {
            Dictionary<string, string> dictAttachmentType = new Dictionary<string, string>
            {
                { "Microsoft Word Document", ".DOC" },
                { "Microsoft Word Open XML Document", ".DOCX" },
                { "Microsoft Excel 97-2003 Worksheet", ".XLS" },
                { "Microsoft Excel workbook", ".XLSX" },
                { "Log File", ".LOG" },
                { "Outlook Mail Message", ".MSG" },
                { "OpenDocument Text Document", ".ODT" },
                { "Pages Document", ".PAGES" },
                { "Rich Text Format File", ".RTF" },
                { "LaTeX Source Document", ".TEX" },
                { "Plain Text File", ".TXT" },
                { "WordPerfect Document", ".WPD" },
                { "Microsoft Works Word Processor Document", ".WPS" },
                { "Comma Separated Values File", ".CSV" },
                { "Data File", ".DAT" },
                { "GEDCOM Genealogy Data File", ".GED" },
                { "Keynote Presentation", ".KEY" },
                { "Mac OS X Keychain File", ".KEYCHAIN" },
                { "PowerPoint Slide Show", ".PPS" },
                { "PowerPoint Presentation", ".PPT" },
                { "PowerPoint Open XML Presentation", ".PPTX" },
                { "Standard Data File", ".SDF" },
                { "Consolidated Unix File Archive", ".TAR" },
                { "TurboTax 2014 Tax Return", ".TAX2014" },
                { "TurboTax 2015 Tax Return", ".TAX2015" },
                { "vCard File", ".VCF" },
                { "XML File", ".XML" },
                { "Audio Interchange File Format", ".AIF" },
                { "Interchange File Format", ".IFF" },
                { "Media Playlist File", ".M3U" },
                { "MPEG-4 Audio File", ".M4A" },
                { "MIDI File", ".MID" },
                { "MP3 Audio File", ".MP3" },
                { "MPEG-2 Audio File", ".MPA" },
                { "WAVE Audio File", ".WAV" },
                { "Windows Media Audio File", ".WMA" },
                { "Rhino 3D Model", ".3DM" },
                { "3D Studio Scene", ".3DS" },
                { "3ds Max Scene File", ".MAX" },
                { "Wavefront 3D Object File", ".OBJ" },
                { "Bitmap Image File", ".BMP" },
                { "DirectDraw Surface", ".DDS" },
                { "Graphical Interchange Format File", ".GIF" },
                { "JPEG Image", ".JPG" },
                { "Portable Network Graphic", ".PNG" },
                { "Adobe Photoshop Document", ".PSD" },
                { "PaintShop Pro Image", ".PSPIMAGE" },
                { "Targa Graphic", ".TGA" },
                { "Thumbnail Image File", ".THM" },
                { "Tagged Image File", ".TIF" },
                { "Tagged Image File Format", ".TIFF" },
                { "YUV Encoded Image File", ".YUV" },
                { "Portable Document Format File", ".PDF" },
                { "Hypertext Markup Language File", ".HTM" },
                //dictAttachmentType.Add("Hypertext Markup Language File", ".HTML");
                { "7-Zip Compressed File", ".7Z" },
                { "Comic Book RAR Archive", ".CBR" },
                { "Debian Software Package", ".DEB" },
                { "Gnu Zipped Archive", ".GZ" },
                { "Mac OS X Installer Package", ".PKG" },
                { "WinRAR Compressed Archive", ".RAR" },
                { "Red Hat Package Manager File", ".RPM" },
                { "StuffIt X Archive", ".SITX" },
                { "Compressed Tarball File", ".TAR.GZ" },
                { "Zipped File", ".ZIP" },
                { "Extended Zip File", ".ZIPX" },
                { "Audio Video Interleave", ".AVI" },
                { "QuickTime", ".MOV" },
                //dictAttachmentType.Add("QuickTime", ".QT");
                { "Advanced Video Coding High Definition", ".AVCHD" },
                { "MPEG Video File", ".MPG" },
                { "MPEG-4 Video File", ".MP4" },
                { "Windows Media Video", ".WMV" },
                { "Matroska Video", ".MKV" },
                { "Musical Instrument Digital Interface", ".MIDI" },
                { "Flash Video", ".FLV" }
            };
            return dictAttachmentType;
        }
        /// <summary>
        /// Language Type dropdown
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeLangaugeTypes()
        {
            var dictLanguageType = new Dictionary<string, string>
            {
                { "Afrikaans", "AF" },
                { "Arabic", "AR" },
                { "Bulgarian", "BG" },
                { "Catalan", "CA" },
                { "Chinese", "ZH" },
                { "Chinese trad.", "ZF" },
                { "Croatian", "HR" },
                { "Customer reserve", "Z1" },
                { "Czech", "CS" },
                { "Danish", "DA" },
                { "Dutch", "NL" },
                { "English", "EN" },
                { "Estonian", "ET" },
                { "Finnish", "FI" },
                { "French", "FR" },
                { "German", "DE" },
                { "Greek", "EL" },
                { "Hebrew", "HE" },
                { "Hungarian", "HU" },
                { "Icelandic", "IS" },
                { "Indonesian", "ID" },
                { "Italian", "IT" },
                { "Japanese", "JA" },
                { "Korean", "KO" },
                { "Latvian", "LV" },
                { "Lithuanian", "LT" },
                { "Malaysian", "MS" },
                { "Norwegian", "NO" },
                { "Polish", "PL" },
                { "Portuguese", "PT" },
                { "Romanian", "RO" },
                { "Russian", "RU" },
                { "Serbian", "SR" },
                { "Serbian (Latin)", "SH" },
                { "Slovakian", "SK" },
                { "Slovenian", "SL" },
                { "Spanish", "ES" },
                { "Swedish", "SV" },
                { "Thai", "TH" },
                { "Turkish", "TR" },
                { "Ukrainian", "UK" }
            };
            return dictLanguageType;
        }
        /// <summary>
        /// Data Type dropdown
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeDataTypes()
        {
            Dictionary<string, string> dictDataType = new Dictionary<string, string>
            {
                { "CHAR: Character String", "CHAR" },
                { "NUMC: Long character", "NUMC" },
                { "DEC: Counter or amount field with decimal point", "DEC" },
                { "ALTN: Alternate Number", "ALTN" },
                { "ISCN: Integration Scenario", "ISCN" },
                { "REQ: Request Type", "REQ" },
                { "STATUS: Status Type", "STATUS" },
                { "EMAIL: Email address", "EMAIL" },
                { "DATS: Date", "DATS" },
                { "DTMS: DateTime in milliseconds", "DTMS" },
                { "TIMS: Time", "TIMS" }
            };
            return dictDataType;
        }
        /// <summary>
        /// dcitionary for data types as per field type selected
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeDataTypesSelection()
        {
            var selectedFieldDataTypes = new Dictionary<string, string>
            {
                { "Text", "Character string, Long character, Counter or amount field with decimal point, Alternate Number, Integration Scenario" },
                { "Dropdown", "Character string, Request Type, Status Type" }
            };

            return selectedFieldDataTypes;
        }
        /// <summary>
        /// Initialize Field Dependencies
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeFieldDependencies()
        {
            Dictionary<string, string> fieldDependencies = new Dictionary<string, string>
            {
                { "Text", "DataType" },
                { "TextArea", "DataType, textAreaLength, textAreaWidth" },
                { "Dropdown", "DataType" },
                { "Module Reference Id", "refObjId" },
                { "Attachment", "AttachmentFileType, AttachmentSize" },
                { "Radio", "optionLimit" }
            };
            return fieldDependencies;
        }
        /// <summary>
        /// Initialize True False Dictionary
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> InitializeTrueFalse()
        {
            Dictionary<string, string> dictTrueFalse = new Dictionary<string, string>
            {
                { "true", "true" },
                { "false", "false" }
            };
            return dictTrueFalse;
        }
        /// <summary>
        /// Initialize the metadata data
        /// </summary>
        /// <returns></returns>
        private static List<FieldMetaDataModel> InitializeMetadataData()
        {
            var metaDataList = new List<FieldMetaDataModel>();
            var metaDataItem = new FieldMetaDataModel() { FieldName = "FieldName", Description = "Field Name", TabName = "Basic Details", HelpText = "Unique Id", IsDropDown = false, DropDownType = "None", Colour = Color.LightGreen, Mandatory = "1" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "FieldDescription", Description = "Field Description", TabName = "Basic Details", HelpText = "Description of Field", IsDropDown = false, DropDownType = "None", Colour = Color.LightGreen, Mandatory = "1" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "FieldType", Description = "Field Type", TabName = "Basic Details", HelpText = "", IsDropDown = true, DropDownType = "Other", Colour = Color.LightGreen, Mandatory = "1" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "DataType", Description = "Data Type", TabName = "Basic Details", HelpText = "", IsDropDown = true, DropDownType = "Other", Colour = Color.LightGreen, Mandatory = "1" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "FieldLength", Description = "Field Length", TabName = "Basic Details", HelpText = "Input maximum length", IsDropDown = false, DropDownType = "None", Colour = Color.LightGreen, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "helpText", Description = "Help Text", TabName = "Basic Details", HelpText = "Additional information for field. It will display on mouse hover of input", IsDropDown = false, DropDownType = "None", Colour = Color.LightGreen, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "outputLength", Description = "Output Length", TabName = "Basic Details", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.LightGreen, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "mandatory", Description = "Mandatory", TabName = "", HelpText = "Marking a field mandatory only works if you have marked the field as a Key field. For non key fields, you can mark the field mandatory while creating the layout of the form.", IsDropDown = true, DropDownType = "YesNo", Colour = Color.DarkGray, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "parentField", Description = "Parent Field", TabName = "", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.DarkGray, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "refObjId", Description = "Reference Object Id", TabName = "Reference Details", HelpText = "", IsDropDown = true, DropDownType = "Other", Colour = Color.Yellow, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "refFieldId", Description = "Reference Field Id", TabName = "Reference Details", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.Yellow, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "searchField", Description = "Search Field", TabName = "Reference Details", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.Yellow, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "isRefField", Description = "Is Reference Field", TabName = "Reference Details", HelpText = "", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.Yellow, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "descFld", Description = "Description Field", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "Keys", Description = "Key Field", TabName = "Other Characteristics", HelpText = "Users are asked to fill the key fields before any other field can be populated. This comes in handy when you want some validations to be done before your user continues to fill the rest of the form", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "permissionFld", Description = "Permission Field", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "numberSettingCrit", Description = "Number Setting Criteria", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "workflowFld", Description = "Workflow Field", TabName = "Other Characteristics", HelpText = "Select this criteria if you envision to configure workflows based on the data of this field. e.g. Decisions and Determination of workflow tasks. The data in this field can also be used for escalation, rejection or any other workflow emails.", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "workflowCriteria", Description = "Workflow Criteria", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "tabCriteriaFld", Description = "Tab Criteria Field", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "srchEngine", Description = "Search Engine", TabName = "Other Characteristics", HelpText = "Selecting this will allow you to assign this field to the list page. This will also allow users to search records based on data stored in this field. It is not recommended for all fields in a module.", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "isTransientField", Description = "Is Transient Field", TabName = "Other Characteristics", HelpText = "While copying a record to create new record, fields marked as transient won't be copied into the new record.", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "suggestion", Description = "Suggestion", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "isCompBased", Description = "Is Company Based", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "isCompleteness", Description = "Is Completeness", TabName = "Other Characteristics", HelpText = "Check to consider the field for completeness", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "checkList", Description = "Check List", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "DecimalPlace", Description = "Decimal Plaes", TabName = "Other Characteristics", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "tnccheck", Description = "Terms N Condition Check", TabName = "Other Characteristics", HelpText = "", IsDropDown = true, DropDownType = "TrueFalse", Colour = Color.LightBlue, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "AttachmentSize", Description = "Attachment Size", TabName = "Attachment Details", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.LightGray, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "AttachmentFileType", Description = "Attachment File Type", TabName = "Attachment Details", HelpText = "", IsDropDown = true, DropDownType = "Other", Colour = Color.LightGray, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "picklistDependency", Description = "Picklist Dependency", TabName = "", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.DarkGray, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "nounFld", Description = "Noun Field", TabName = "", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.DarkGray, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "optionLimit", Description = "Option Limit", TabName = "Radio Button Properties", HelpText = "", IsDropDown = false, DropDownType = "None", Colour = Color.LightSalmon, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "textAreaLength", Description = "Textarea Length", TabName = "Textarea Properties", HelpText = "For the length of textarea", IsDropDown = false, DropDownType = "None", Colour = Color.LightCoral, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "textAreaWidth", Description = "Textarea Width", TabName = "Textarea Properties", HelpText = "For the width of textarea", IsDropDown = false, DropDownType = "None", Colour = Color.LightCoral, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "defaultDate", Description = "Default Date", TabName = "Date Properties", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightCyan, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "futureDate", Description = "Future Date", TabName = "Date Properties", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightCyan, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "pastDate", Description = "Past Date", TabName = "Date Properties", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightCyan, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "gridDisplay", Description = "Grid Display", TabName = "Table Properties", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightGoldenrodYellow, Mandatory = "0" };
            metaDataList.Add(metaDataItem);
            metaDataItem = new FieldMetaDataModel() { FieldName = "defDisplay", Description = "Default Display", TabName = "Table Properties", HelpText = "", IsDropDown = true, DropDownType = "YesNo", Colour = Color.LightGoldenrodYellow, Mandatory = "0" };
            metaDataList.Add(metaDataItem);

            return metaDataList;
        }
        /// <summary>
        /// Initialize the dropdown header
        /// </summary>
        /// <returns></returns>
        private static List<DropdownHeadersModel> InitializeDropdownHeadersModel()
        {
            var dropDownHeadersList = new List<DropdownHeadersModel>();
            var dropDownHeadersItem = new DropdownHeadersModel { FieldName = "FieldId", Description = "Field Id", IsDropDown = false };
            dropDownHeadersList.Add(dropDownHeadersItem);
            dropDownHeadersItem = new DropdownHeadersModel { FieldName = "Code", Description = "Code", IsDropDown = false };
            dropDownHeadersList.Add(dropDownHeadersItem);
            dropDownHeadersItem = new DropdownHeadersModel { FieldName = "Text", Description = "Text", IsDropDown = false };
            dropDownHeadersList.Add(dropDownHeadersItem);
            dropDownHeadersItem = new DropdownHeadersModel { FieldName = "Language", Description = "Language", IsDropDown = true };
            dropDownHeadersList.Add(dropDownHeadersItem);
            dropDownHeadersItem = new DropdownHeadersModel { FieldName = "ParentId", Description = "Parent Id", IsDropDown = false };
            dropDownHeadersList.Add(dropDownHeadersItem);
            dropDownHeadersItem = new DropdownHeadersModel { FieldName = "ParentValue", Description = "Parent Value", IsDropDown = false };
            dropDownHeadersList.Add(dropDownHeadersItem);
            dropDownHeadersItem = new DropdownHeadersModel { FieldName = "PlantCode", Description = "Plant Code", IsDropDown = false };
            dropDownHeadersList.Add(dropDownHeadersItem);
            return dropDownHeadersList;
        }
        /// <summary>
        /// Initialize the translation headers
        /// </summary>
        /// <returns></returns>
        private static List<TranslationHeader> InitializeTranslationHeaders()
        {
            var translationHeader = new List<TranslationHeader>();
            var translationHeaderItem = new TranslationHeader { FieldName = "FieldId", Description = "Field Id", IsDropDown = false };
            translationHeader.Add(translationHeaderItem);
            translationHeaderItem = new TranslationHeader { FieldName = "FieldDesc", Description = "Description (English)", IsDropDown = false };
            translationHeader.Add(translationHeaderItem);
            translationHeaderItem = new TranslationHeader { FieldName = "Language", Description = "Language", IsDropDown = true };
            translationHeader.Add(translationHeaderItem);
            translationHeaderItem = new TranslationHeader { FieldName = "ShortText", Description = "Short Text", IsDropDown = false };
            translationHeader.Add(translationHeaderItem);
            translationHeaderItem = new TranslationHeader { FieldName = "LongText", Description = "Long Text", IsDropDown = false };
            translationHeader.Add(translationHeaderItem);
            return translationHeader;
        }
        /// <summary>
        /// Initialize Sql KeyWords
        /// </summary>
        /// <returns></returns>
        private static List<string> InitializeSqlKeyWords()
        {
            List<string> lstSqlKeyWords = new List<string>
            {
                "ADD",
                "ALL",
                "ALTER",
                "AND",
                "ANY",
                "AS",
                "ASC",
                "AUTHORIZATION",
                "BACKUP",
                "BEGIN",
                "BETWEEN",
                "BREAK",
                "BROWSE",
                "BULK",
                "BY",
                "CASCADE",
                "CASE",
                "CHECK",
                "CHECKPOINT",
                "CLOSE",
                "CLUSTERED",
                "COALESCE",
                "COLLATE",
                "COLUMN",
                "COMMIT",
                "COMPUTE",
                "CONSTRAINT",
                "CONTAINS",
                "CONTAINSTABLE",
                "CONTINUE",
                "CONVERT",
                "CREATE",
                "CROSS",
                "CURRENT",
                "CURRENT_DATE",
                "CURRENT_TIME",
                "CURRENT_TIMESTAMP",
                "CURRENT_USER",
                "CURSOR",
                "DATABASE",
                "DBCC",
                "DEALLOCATE",
                "DECLARE",
                "DEFAULT",
                "DELETE",
                "DENY",
                "DESC",
                "DISK",
                "DISTINCT",
                "DISTRIBUTED",
                "DOUBLE",
                "DROP",
                "DUMMY",
                "DUMP",
                "ELSE",
                "END",
                "ERRLVL",
                "ESCAPE",
                "ABSOLUTE",
                "ACTION",
                "ADMIN",
                "AFTER",
                "AGGREGATE",
                "ALIAS",
                "ALLOCATE",
                "ARE",
                "ARRAY",
                "ASSERTION",
                "AT",
                "BEFORE",
                "BINARY",
                "BIT",
                "BLOB",
                "BOOLEAN",
                "BOTH",
                "BREADTH",
                "CALL",
                "CASCADED",
                "CAST",
                "CATALOG",
                "CHAR",
                "CHARACTER",
                "CLASS",
                "CLOB",
                "COLLATION",
                "COMPLETION",
                "CONNECT",
                "CONNECTION",
                "CONSTRAINTS",
                "CONSTRUCTOR",
                "CORRESPONDING",
                "CUBE",
                "CURRENT_PATH",
                "CURRENT_ROLE",
                "CYCLE",
                "DATA",
                "DATE",
                "DAY",
                "DEC",
                "DECIMAL",
                "DEFERRABLE",
                "DEFERRED",
                "DEPTH",
                "DEREF",
                "DESCRIBE",
                "DESCRIPTOR",
                "DESTROY",
                "DESTRUCTOR",
                "DETERMINISTIC",
                "DICTIONARY",
                "DIAGNOSTICS",
                "DISCONNECT",
                "DOMAIN",
                "DYNAMIC",
                "EACH",
                "END-EXEC",
                "EQUALS",
                "EVERY",
                "EXCEPTION",
                "EXTERNAL",
                "FALSE",
                "FIRST",
                "FLOAT",
                "EXCEPT",
                "EXEC",
                "EXECUTE",
                "EXISTS",
                "EXIT",
                "FETCH",
                "FILE",
                "FILLFACTOR",
                "FOR",
                "FOREIGN",
                "FREETEXT",
                "FREETEXTTABLE",
                "FROM",
                "FULL",
                "FUNCTION",
                "GOTO",
                "GRANT",
                "GROUP",
                "HAVING",
                "HOLDLOCK",
                "IDENTITY",
                "IDENTITY_INSERT",
                "IDENTITYCOL",
                "IF",
                "IN",
                "INDEX",
                "INNER",
                "INSERT",
                "INTERSECT",
                "INTO",
                "IS",
                "JOIN",
                "KEY",
                "KILL",
                "LEFT",
                "LIKE",
                "LINENO",
                "LOAD",
                "NATIONAL",
                "NOCHECK",
                "NONCLUSTERED",
                "NOT",
                "NULL",
                "NULLIF",
                "OF",
                "OFF",
                "OFFSETS",
                "ON",
                "OPEN",
                "OPENDATASOURCE",
                "OPENQUERY",
                "OPENROWSET",
                "OPENXML",
                "OPTION",
                "OR",
                "ORDER",
                "OUTER",
                "OVER",
                "FOUND",
                "FREE",
                "GENERAL",
                "GET",
                "GLOBAL",
                "GO",
                "GROUPING",
                "HOST",
                "HOUR",
                "IGNORE",
                "IMMEDIATE",
                "INDICATOR",
                "INITIALIZE",
                "INITIALLY",
                "INOUT",
                "INPUT",
                "INT",
                "INTEGER",
                "INTERVAL",
                "ISOLATION",
                "ITERATE",
                "LANGUAGE",
                "LARGE",
                "LAST",
                "LATERAL",
                "LEADING",
                "LESS",
                "LEVEL",
                "LIMIT",
                "LOCAL",
                "LOCALTIME",
                "LOCALTIMESTAMP",
                "LOCATOR",
                "MAP",
                "MATCH",
                "MINUTE",
                "MODIFIES",
                "MODIFY",
                "MODULE",
                "MONTH",
                "NAMES",
                "NATURAL",
                "NCHAR",
                "NCLOB",
                "NEW",
                "NEXT",
                "NO",
                "NONE",
                "NUMERIC",
                "OBJECT",
                "OLD",
                "ONLY",
                "OPERATION",
                "ORDINALITY",
                "OUT",
                "OUTPUT",
                "PAD",
                "PARAMETER",
                "PARAMETERS",
                "PARTIAL",
                "PATH",
                "POSTFIX",
                "PREFIX",
                "PREORDER",
                "PREPARE",
                "PERCENT",
                "PLAN",
                "PRECISION",
                "PRIMARY",
                "PRINT",
                "PROC",
                "PROCEDURE",
                "PUBLIC",
                "RAISERROR",
                "READ",
                "READTEXT",
                "RECONFIGURE",
                "REFERENCES",
                "REPLICATION",
                "RESTORE",
                "RESTRICT",
                "RETURN",
                "REVOKE",
                "RIGHT",
                "ROLLBACK",
                "ROWCOUNT",
                "ROWGUIDCOL",
                "RULE",
                "SAVE",
                "SCHEMA",
                "SELECT",
                "SESSION_USER",
                "SET",
                "SETUSER",
                "SHUTDOWN",
                "SOME",
                "STATISTICS",
                "SYSTEM_USER",
                "TABLE",
                "TEXTSIZE",
                "THEN",
                "TO",
                "TOP",
                "TRAN",
                "TRANSACTION",
                "TRIGGER",
                "TRUNCATE",
                "TSEQUAL",
                "UNION",
                "UNIQUE",
                "UPDATE",
                "UPDATETEXT",
                "USE",
                "USER",
                "VALUES",
                "VARYING",
                "VIEW",
                "WAITFOR",
                "WHEN",
                "WHERE",
                "WHILE",
                "WITH",
                "WRITETEXT",
                "PRESERVE",
                "PRIOR",
                "PRIVILEGES",
                "READS",
                "REAL",
                "RECURSIVE",
                "REF",
                "REFERENCING",
                "RELATIVE",
                "RESULT",
                "RETURNS",
                "ROLE",
                "ROLLUP",
                "ROUTINE",
                "ROW",
                "ROWS",
                "SAVEPOINT",
                "SCROLL",
                "SCOPE",
                "SEARCH",
                "SECOND",
                "SECTION",
                "SEQUENCE",
                "SESSION",
                "SETS",
                "SIZE",
                "SMALLINT",
                "SPACE",
                "SPECIFIC",
                "SPECIFICTYPE",
                "SQL",
                "SQLEXCEPTION",
                "SQLSTATE",
                "SQLWARNING",
                "START",
                "STATE",
                "STATEMENT",
                "STATIC",
                "STRUCTURE",
                "TEMPORARY",
                "TERMINATE",
                "THAN",
                "TIME",
                "TIMESTAMP",
                "TIMEZONE_HOUR",
                "TIMEZONE_MINUTE",
                "TRAILING",
                "TRANSLATION",
                "TREAT",
                "TRUE",
                "UNDER",
                "UNKNOWN",
                "UNNEST",
                "USAGE",
                "USING",
                "VALUE",
                "VARCHAR",
                "VARIABLE",
                "WHENEVER",
                "WITHOUT",
                "WORK",
                "WRITE",
                "YEAR",
                "ZONE"
            };
            return lstSqlKeyWords;
        }
        /// <summary>
        /// Method to populate the header on the dropdown sheet
        /// </summary>
        /// <returns></returns>
        private static List<string> InitializeDropDownHeaders()
        {
            List<string> dropDownHeader = new List<string>
            {
                "CODE",
                "TEXT",
                "LANGUAGE"
            };
            return dropDownHeader;
        }
        /// <summary>
        /// Initialize default field lengths
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, int> InitialiazeDefaultFieldLength()
        {
            Dictionary<string, int> dictDefaultFieldLength = new Dictionary<string, int>
            {
                { "Text", 100 },
                { "Dropdown", 100 },
                { "Checkbox", 5 },
                { "Radio", 100 },
                { "TextArea", 100 },
                { "Date", 20 },
                { "Time", 20 },
                { "Date and Time", 100 },
                { "Module Reference Id", 100 },
                { "Grid", 100 }
            };
            return dictDefaultFieldLength;
        }
        #endregion
        #region Public Methods
        /// <summary>
        /// Method used to initialize the class variables
        /// </summary>
        public static void InitializeClassVariables()
        {
            GlobalMembers.InstanceGlobalMembers.ListMetadataData = InitializeMetadataData();
            GlobalMembers.InstanceGlobalMembers.DictionaryFieldTypes = InitializeFieldTypes();
            GlobalMembers.InstanceGlobalMembers.DictionaryDataTypes = InitializeDataTypes();
            GlobalMembers.InstanceGlobalMembers.DictionaryAttachmentFileTypes = InitializeAttachmentTypes();
            GlobalMembers.InstanceGlobalMembers.DictionaryYesNo = InitializeYesNo();
            GlobalMembers.InstanceGlobalMembers.DictionaryTrueFalse = InitializeTrueFalse();
            GlobalMembers.InstanceGlobalMembers.DictionaryFieldDependencies = InitializeFieldDependencies();
            GlobalMembers.InstanceGlobalMembers.DictionaryDefaultFieldLength = InitialiazeDefaultFieldLength();
            GlobalMembers.InstanceGlobalMembers.ListSqlKeyWords = InitializeSqlKeyWords();
            GlobalMembers.InstanceGlobalMembers.ListDropdownHeaders = InitializeDropDownHeaders();
            GlobalMembers.InstanceGlobalMembers.ListDropdownHeadersModels = InitializeDropdownHeadersModel();
            GlobalMembers.InstanceGlobalMembers.DictionaryLanguageType = InitializeLangaugeTypes();
            GlobalMembers.InstanceGlobalMembers.ListTranslationHeaders = InitializeTranslationHeaders();
        }
        /// <summary>
        /// Method used to check if the tables exists or not. If not then create the table and populate them with data
        /// </summary>
        public static void InitializeDatabaseTables()
        {
            if (GlobalMembers.InstanceGlobalMembers.DtMetadataData.Rows.Count == 0)
            { GlobalMembers.InstanceGlobalMembers.DtMetadataData = ExcelHelper.PopulateMetadataData(); }
            if (GlobalMembers.InstanceGlobalMembers.DtAttachmentTypeValues.Rows.Count == 0)
            { ExcelHelper.PopulateDropDowns("AttachmentTypeValues"); }
            if (GlobalMembers.InstanceGlobalMembers.DtDataTypeValues.Rows.Count == 0)
            { ExcelHelper.PopulateDropDowns("DataTypeValues"); }
            if (GlobalMembers.InstanceGlobalMembers.DtFieldTypeValues.Rows.Count == 0)
            { ExcelHelper.PopulateDropDowns("FieldTypeValues"); }
            if (GlobalMembers.InstanceGlobalMembers.DtTrueFalseValues.Rows.Count == 0)
            { ExcelHelper.PopulateDropDowns("TrueFalseValues"); }
            if (GlobalMembers.InstanceGlobalMembers.DtYesNoValues.Rows.Count == 0)
            { ExcelHelper.PopulateDropDowns("YesNoValues"); }
        }
        #endregion
        #region Old Methods For Reference
        private static Dictionary<string, int> InitialiazeDefaultFieldLength_Old()
        {
            Dictionary<string, int> dictDefaultFieldLength = new Dictionary<string, int>
            {
                { "Text", 100 },
                { "Dropdown", 100 },
                { "Checkbox", 5 },
                { "Radio", 100 },
                { "TextArea", 100 },
                { "Date", 20 },
                { "Time", 20 },
                { "Date and Time", 100 },
                { "Module Reference Id", 100 },
                { "Grid", 100 }
            };

            /* To be used later (Don't delete)
            dictDefaultFieldLength.Add("User Selection", 100);
            dictDefaultFieldLength.Add("Location Reference", 100);
            dictDefaultFieldLength.Add("Attachment", 100);
            dictDefaultFieldLength.Add("Digital Sign", 100);
            dictDefaultFieldLength.Add("Aggregation", 100);
            dictDefaultFieldLength.Add("Translation Table", 100);
            dictDefaultFieldLength.Add("HTML", 100);
            dictDefaultFieldLength.Add("Calculation", 100);
            dictDefaultFieldLength.Add("URL", 100);
            dictDefaultFieldLength.Add("Position Type", 100);
            dictDefaultFieldLength.Add("Tree", 100);
            dictDefaultFieldLength.Add("Noun Type", 100);
            dictDefaultFieldLength.Add("Work Flow", 100);
            dictDefaultFieldLength.Add("Rejection Type", 100);
            dictDefaultFieldLength.Add("Activate/ Deactivate", 100);
            dictDefaultFieldLength.Add("GeoLocation", 100);
            dictDefaultFieldLength.Add("Group", 100);
            dictDefaultFieldLength.Add("Email", 100);
            dictDefaultFieldLength.Add("Password", 100);
            */
            return dictDefaultFieldLength;
        }
        private static Dictionary<string, string> InitializeDataTypes_old()
        {
            Dictionary<string, string> dictDataType = new Dictionary<string, string>
            {
                //dictDataType.Add("Alternate Number", "ALTN");
                //dictDataType.Add("Character string", "CHAR");
                //dictDataType.Add("Date", "DATS");
                //dictDataType.Add("Counter or amount field with decimal point", "DEC");
                //dictDataType.Add("DateTime in milliseconds", "DTMS");
                //dictDataType.Add("Email address", "EMAIL");
                //dictDataType.Add("Floating-point number", "FLTP");
                //dictDataType.Add("1-byte", "INT1");
                //dictDataType.Add("2-byte", "INT2");
                //dictDataType.Add("4-byte", "INT4");
                //dictDataType.Add("Integration Scenario", "ISCN");
                //dictDataType.Add("Language key", "LANG");
                //dictDataType.Add("Character string", "LCHR");
                //dictDataType.Add("Uninterpreted byte string", "LRAW");
                //dictDataType.Add("Long character", "NUMC");
                //dictDataType.Add("Password", "PASS");
                //dictDataType.Add("Accuracy of a QUAN", "PREC");
                //dictDataType.Add("Quantity", "QUAN");
                //dictDataType.Add("Uninterpreted sequence of bytes", "RAW");
                //dictDataType.Add("Request Type", "REQ");
                //dictDataType.Add("Status Type", "STATUS");
                //dictDataType.Add("Time", "TIMS");
                //dictDataType.Add("Units key", "UNIT");
                //dictDataType.Add("Character field of variable length", "VARC");

                { "CHAR: Character String", "CHAR" },
                { "NUMC: Long character", "NUMC" },
                { "DEC: Counter or amount field with decimal point", "DEC" },
                { "ALTN: Alternate Number", "ALTN" },
                { "ISCN: Integration Scenario", "ISCN" },
                { "REQ: Request Type", "REQ" },
                { "STATUS: Status Type", "STATUS" },
                { "EMAIL: Email address", "EMAIL" },
                { "DATS: Date", "DATS" },
                { "DTMS: DateTime in milliseconds", "DTMS" },
                { "TIMS: Time", "TIMS" }
            };
            return dictDataType;
        }
        #endregion
    }
}
