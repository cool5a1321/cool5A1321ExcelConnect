using Autofac;
using ConnektHub.Models;
using Prospecta.ConnektHub.SQLiteHelper;
using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prospecta.ConnektHub.Models
{
    public class GlobalMembers
    {
        #region private members
        private static GlobalMembers globalMembers = new GlobalMembers();
        #endregion
        #region public properties
        public static GlobalMembers InstanceGlobalMembers
        { get { return globalMembers; } }
        #endregion
        #region constructors
        private GlobalMembers()
        { }
        #endregion
        #region public members
        public Excel.Application ExcelApplication;
        

        public string IntroductionSheetName;
        public string ConfigurationSheetName;
        public string MetaDataSheetName;
        public string DropDownSheetName;
        public string TranslationSheetName;

        public SQLiteDatabase SqliteDatabase;

        public DataTable DtMetadataData;
        public DataTable DtFieldTypeValues;
        public DataTable DtDataTypeValues;
        public DataTable DtAttachmentTypeValues;
        public DataTable DtTrueFalseValues;
        public DataTable DtYesNoValues;

        public List<FieldMetaDataModel> ListMetadataData;
        public Dictionary<string, string> DictionaryModulesList;
        public Dictionary<string, string> DictionaryDataTypes;
        public Dictionary<string, string> DictionaryFieldTypes;
        public Dictionary<string, string> DictionaryAttachmentFileTypes;
        public Dictionary<string, string> DictionaryYesNo;
        public Dictionary<string, string> DictionaryTrueFalse;
        public Dictionary<string, string> DictionaryFieldDependencies;
        public Dictionary<string, int> DictionaryDefaultFieldLength;
        public Dictionary<string, string> DictionaryLanguageType;
        public List<string> ListDropdownHeaders;
        public List<string> ListSqlKeyWords;
        public List<TranslationHeader> ListTranslationHeaders;
        public List<DropdownHeadersModel> ListDropdownHeadersModels;

        public IContainer Container = null;
        #endregion
    }
}