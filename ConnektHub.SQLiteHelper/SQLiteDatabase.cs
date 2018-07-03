using Prospecta.ConnektHub.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Prospecta.ConnektHub.SQLiteHelper
{
    public class SQLiteDatabase : IDisposable
    {
        readonly String DBConnection;

        private readonly SQLiteTransaction _sqLiteTransaction;

        private readonly SQLiteConnection _sqLiteConnection;

        private readonly bool _transaction;

        /// <summary>
        ///     Default Constructor for SQLiteDatabase Class.
        /// </summary>
        /// <param name="transaction">Allow programmers to insert, update and delete values in one transaction</param>
        public SQLiteDatabase(bool transaction = false)
        {
            _transaction = transaction;
            DBConnection = "Data Source=recipes.s3db";
            if (transaction)
            {
                _sqLiteConnection = new SQLiteConnection(DBConnection);
                _sqLiteConnection.Open();
                _sqLiteTransaction = _sqLiteConnection.BeginTransaction();
            }
        }

        /// <summary>
        ///     Single Param Constructor for specifying the DB file.
        /// </summary>
        /// <param name="inputFile">The File containing the DB</param>
        public SQLiteDatabase(String inputFile)
        {
            if (!File.Exists(inputFile))
            { CreateDB(inputFile); }
            DBConnection = String.Format("Data Source={0}", inputFile);
        }

        /// <summary>
        ///     Commit transaction to the database.
        /// </summary>
        public void CommitTransaction()
        {
            _sqLiteTransaction.Commit();
            _sqLiteTransaction.Dispose();
            _sqLiteConnection.Close();
            _sqLiteConnection.Dispose();
        }

        /// <summary>
        ///     Single Param Constructor for specifying advanced connection options.
        /// </summary>
        /// <param name="connectionOpts">A dictionary containing all desired options and their values</param>
        public SQLiteDatabase(Dictionary<String, String> connectionOpts)
        {
            String str = connectionOpts.Aggregate("", (current, row) => current + String.Format("{0}={1}; ", row.Key, row.Value));
            str = str.Trim().Substring(0, str.Length - 1);
            DBConnection = str;
        }

        /// <summary>
        ///     Allows the programmer to create new database file.
        /// </summary>
        /// <param name="filePath">Full path of a new database file.</param>
        /// <returns>true or false to represent success or failure.</returns>
        public bool CreateDB(string filePath)
        {
            try
            {
                SQLiteConnection.CreateFile(filePath);
                return true;
            }
#pragma warning disable CS0168 // Variable is declared but never used
            catch (Exception e)
#pragma warning restore CS0168 // Variable is declared but never used
            {
                return false;
            }
        }

        public DataTable GetDataTable(string sql)
        {
            var dt = new DataTable();
            try
            {
                SQLiteDataAdapter sQLiteDataAdapter;
                using (var conn = new SQLiteConnection(DBConnection))
                {
                    conn.Open();
                    using (sQLiteDataAdapter = new SQLiteDataAdapter(sql, conn))
                    {
                        sQLiteDataAdapter.Fill(dt);
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            { throw (ex); }
            return dt;
        }

        /// <summary>
        ///     Allows the programmer to run a query against the Database.
        /// </summary>
        /// <param name="sql">The SQL to run</param>
        /// <param name="allowDBNullColumns">Allow null value for columns in this collection.</param>
        /// <returns>A DataTable containing the result set.</returns>
        public DataTable GetDataTable(string sql, IEnumerable<string> allowDBNullColumns = null)
        {
            var dt = new DataTable();
            if (allowDBNullColumns != null)
                foreach (var s in allowDBNullColumns)
                {
                    dt.Columns.Add(s);
                    dt.Columns[s].AllowDBNull = true;
                }
            try
            {
                var cnn = new SQLiteConnection(DBConnection);
                cnn.Open();
                var mycommand = new SQLiteCommand(cnn) { CommandText = sql };
                var reader = mycommand.ExecuteReader();
                dt.Load(reader);
                reader.Close();
                cnn.Close();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return dt;
        }

        public string RetrieveOriginal(string value)
        {
            return
                value.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", "<").Replace("&quot;", "\"").Replace(
                    "&apos;", "'");
        }

        /// <summary>
        ///     Allows the programmer to interact with the database for purposes other than a query.
        /// </summary>
        /// <param name="sql">The SQL to be run.</param>
        /// <returns>An Integer containing the number of rows updated.</returns>
        public int ExecuteNonQuery(string sql)
        {
            if (!_transaction)
            {
                var cnn = new SQLiteConnection(DBConnection);
                cnn.Open();
                var mycommand = new SQLiteCommand(cnn) { CommandText = sql };
                var rowsUpdated = mycommand.ExecuteNonQuery();
                cnn.Close();
                return rowsUpdated;
            }
            else
            {
                var mycommand = new SQLiteCommand(_sqLiteConnection) { CommandText = sql };
                return mycommand.ExecuteNonQuery();
            }
        }

        /// <summary>
        ///     Allows the programmer to retrieve single items from the DB.
        /// </summary>
        /// <param name="sql">The query to run.</param>
        /// <returns>A string.</returns>
        public string ExecuteScalar(string sql)
        {
            if (!_transaction)
            {
                var cnn = new SQLiteConnection(DBConnection);
                cnn.Open();
                var mycommand = new SQLiteCommand(cnn) { CommandText = sql };
                var value = mycommand.ExecuteScalar();
                cnn.Close();
                return value != null ? value.ToString() : "";
            }
            else
            {
                var sqLiteCommand = new SQLiteCommand(_sqLiteConnection) { CommandText = sql };
                var value = sqLiteCommand.ExecuteScalar();
                return value != null ? value.ToString() : "";
            }
        }

        /// <summary>
        ///     Allows the programmer to easily update rows in the DB.
        /// </summary>
        /// <param name="tableName">The table to update.</param>
        /// <param name="data">A dictionary containing Column names and their new values.</param>
        /// <param name="where">The where clause for the update statement.</param>
        /// <returns>A boolean true or false to signify success or failure.</returns>
        public bool Update(String tableName, Dictionary<String, String> data, String where)
        {
            String vals = "";
            Boolean returnCode = true;
            if (data.Count >= 1)
            {
                vals = data.Aggregate(vals, (current, val) => current + String.Format(" {0} = '{1}',", val.Key.ToString(CultureInfo.InvariantCulture), val.Value.ToString(CultureInfo.InvariantCulture)));
                vals = vals.Substring(0, vals.Length - 1);
            }
            try
            {
                ExecuteNonQuery(String.Format("update {0} set {1} where {2};", tableName, vals, where));
            }
            catch
            {
                returnCode = false;
            }
            return returnCode;
        }

        /// <summary>
        ///     Allows the programmer to easily delete rows from the DB.
        /// </summary>
        /// <param name="tableName">The table from which to delete.</param>
        /// <param name="where">The where clause for the delete.</param>
        /// <returns>A boolean true or false to signify success or failure.</returns>
        public bool Delete(string tableName, string where)
        {
            Boolean returnCode = true;
            try
            {
                if (where.Equals(string.Empty))
                { ExecuteNonQuery(string.Format("delete from {0};", tableName)); }
                else
                { ExecuteNonQuery(string.Format("delete from {0} where {1};", tableName, where)); }
                
            }
#pragma warning disable CS0168 // Variable is declared but never used
            catch (Exception fail)
#pragma warning disable CS0168 // Variable is declared but never used
            {
                returnCode = false;
            }
            return returnCode;
        }

        /// <summary>
        ///     Allows the programmer to easily insert into the DB
        /// </summary>
        /// <param name="tableName">The table into which we insert the data.</param>
        /// <param name="data">A dictionary containing the column names and data for the insert.</param>
        /// <returns>returns last inserted row id if it's value is zero than it means failure.</returns>
        public long Insert(String tableName, Dictionary<String, String> data)
        {
            String columns = string.Empty, values = string.Empty, insertCommand = string.Empty, returnValue = string.Empty;
            
            foreach (KeyValuePair<String, String> val in data)
            {
                columns += String.Format(" {0},", val.Key.ToString(CultureInfo.InvariantCulture));
                values += String.Format(" '{0}',", val.Value.Replace("'", "''"));
            }
            columns = columns.Substring(0, columns.Length - 1);
            values = values.Substring(0, values.Length - 1);

            using (var conn = new SQLiteConnection(DBConnection))
            {
                conn.Open();

                using (var command = new SQLiteCommand(conn))
                {
                    command.CommandText = "SELECT name FROM sqlite_master WHERE name='" + tableName + "' and type = 'table'";
                    var name = command.ExecuteScalar();

                    if (name == null)
                    {
                        //command.CommandText = "CREATE TABLE " + tableName + " (CODE VARCHAR(20), TEXT VARCHAR(20), PARENT_FIELD VARCHAR(20))";
                        command.CommandText = "CREATE TABLE " + tableName + " (" + columns.Replace(",", " VARCHAR(20),") + " VARCHAR(20)" + ")";
                        command.ExecuteNonQuery();
                    }

                    using (var transaction = conn.BeginTransaction())
                    {
                        insertCommand = string.Format("insert into {0}({1}) values({2});", tableName, columns, values);
                        var sqLiteCommand = new SQLiteCommand(conn) { CommandText = insertCommand };
                        sqLiteCommand.ExecuteNonQuery();

                        sqLiteCommand = new SQLiteCommand(conn) { CommandText = "SELECT last_insert_rowid()" };
                        returnValue = sqLiteCommand.ExecuteScalar().ToString();
                        transaction.Commit();
                    }
                }

                conn.Close();
            }
            return long.Parse(returnValue);
        }

        public void BulkInsert(ObservableCollection<DropDown> dropValues, string tableName)
        {

            using (var conn = new SQLiteConnection(DBConnection))
            {
                conn.Open();

                using (var command = new SQLiteCommand(conn))
                {
                    command.CommandText = "SELECT name FROM sqlite_master WHERE name=" + tableName;
                    var name = command.ExecuteScalar();

                    if (name == null)
                    {
                        command.CommandText = "CREATE TABLE " + tableName + " (CODE VARCHAR(20), TEXT VARCHAR(20), PARENT_FIELD VARCHAR(20))";
                        command.ExecuteNonQuery();
                    }

                    using (var transaction = conn.BeginTransaction())
                    {
                        foreach (var dropItem in dropValues)
                        {
                            command.CommandText = "INSERT INTO " + tableName + " (CODE, TEXT) VALUES (@dropCode, @dropText);";
                            command.Parameters.Add("@dropCode", DbType.VarNumeric).Value = dropItem.CODE;
                            command.Parameters.Add("@dropText", DbType.VarNumeric).Value = dropItem.TEXT;

                            command.ExecuteNonQuery();
                        }

                        transaction.Commit();
                    }
                }

                conn.Close();
            }
        }
        /// <summary>
        ///     Allows the programmer to easily delete all data from the DB.
        /// </summary>
        /// <returns>A boolean true or false to signify success or failure.</returns>
        public bool ClearDB()
        {
            try
            {
                var tables = GetDataTable("select NAME from SQLITE_MASTER where type='table' order by NAME;");
                foreach (DataRow table in tables.Rows)
                {
                    ClearTable(table["NAME"].ToString());
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        ///     Allows the user to easily clear all data from a specific table.
        /// </summary>
        /// <param name="table">The name of the table to clear.</param>
        /// <returns>A boolean true or false to signify success or failure.</returns>
        public bool ClearTable(String table)
        {
            try
            {
                ExecuteNonQuery(String.Format("delete from {0};", table));
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        ///     Allows the user to easily reduce size of database.
        /// </summary>
        /// <returns>A boolean true or false to signify success or failure.</returns>
        public bool CompactDB()
        {
            try
            {
                ExecuteNonQuery("Vacuum;");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        /// <summary>
        /// Check if the table exists or not
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public bool CheckIfTableExists(string tableName)
        {
            bool tableExists = false;
            using (var conn = new SQLiteConnection(DBConnection))
            {
                conn.Open();

                using (var command = new SQLiteCommand(conn))
                {
                    command.CommandText = "SELECT name FROM sqlite_master WHERE name='" + tableName + "' and type = 'table'";
                    var name = command.ExecuteScalar();
                    if (name != null)
                    { tableExists = true; }
                }
            }
            return tableExists;
        }
        /// <summary>
        /// Method is used to create a table
        /// </summary>
        /// <param name="tableName"></param>
        public void CreateTable(string tableName)
        {
            using (var conn = new SQLiteConnection(DBConnection))
            {
                conn.Open();
                using (var command = new SQLiteCommand(conn))
                {
                    command.CommandText = "CREATE TABLE " + tableName + " (Code TEXT, Value TEXT, PRIMARY KEY(Code))";
                    command.ExecuteNonQuery();
                }
            }
        }

        public void CreateTable(string tableName, string[] columnsNames, string[] dataTypes)
        {
            string createCommand = string.Empty, temp = string.Empty;
            using (var conn = new SQLiteConnection(DBConnection))
            {
                conn.Open();
                using (var command = new SQLiteCommand(conn))
                {
                    
                    if (columnsNames.Length == dataTypes.Length)
                    {
                        createCommand = "CREATE TABLE " + tableName + "(";
                        for (int i = 0; i < columnsNames.Length; i++)
                        {
                            if (string.IsNullOrEmpty(temp))
                            { temp = columnsNames[i] + " " + dataTypes[i]; }
                            else
                            { temp += "," + columnsNames[i] + " " + dataTypes[i]; }
                        }
                        createCommand += temp;
                        createCommand += ")";
                        command.CommandText = createCommand;
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
        /// <summary>
        /// Dispose method
        /// </summary>
        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}