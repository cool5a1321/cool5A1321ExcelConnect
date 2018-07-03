using Microsoft.Office.Interop.Excel;
using System;

namespace Prospecta.ConnektHub.Helpers
{
    public class HelperUtil
    {
        /// <summary>
        /// This method returns the 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="sh"></param>
        /// <returns></returns>
        public static Worksheet GetSheetNameFromGroupOfSheets(string sheetName, Sheets sh)
        {
            try
            {
                foreach (Worksheet worksheet in sh)
                {
                    if (worksheet.Name == sheetName)
                    {
                        return worksheet;
                    }
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            return null;
        }

        /// <summary>
        /// Check if all the characters are in upper case or not
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool IsAllUpper(string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                if (Char.IsLetter(input[i]) && !Char.IsUpper(input[i]))
                    return false;
            }
            return true;
        }
    }
}
