using Prospecta.ConnektHub.Core;
using System;
using System.Collections.Generic;
using System.IO;

namespace Prospecta.ConnektHub.SQLiteHelper.User
{
    public class SQLiteUser
    {
        public static bool AddUserDetails(UserDetails userDetails)
        {
            SQLiteDatabase sQLiteDatabase = new SQLiteDatabase();


            string dbFileLocation = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dbLocation = Path.Combine(dbFileLocation, "Testing2.s3db");
            if (!File.Exists(dbLocation))
            {
                sQLiteDatabase.CreateDB(dbLocation);
            }

            sQLiteDatabase = new SQLiteDatabase(dbLocation);
            Dictionary<string, string> dictLogin = new Dictionary<string, string>();
            dictLogin.Add("username", userDetails.userName);
            dictLogin.Add("password", userDetails.password);
            dictLogin.Add("fullname", userDetails.fullName);
            var retVal = sQLiteDatabase.Insert("UserDetails", dictLogin);

            return (retVal > 0);
        }
    }
}
