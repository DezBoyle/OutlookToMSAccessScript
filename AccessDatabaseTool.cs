using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Net;
using System.Runtime.Versioning;

namespace OutlookToMSAccessScript
{
    [SupportedOSPlatform("windows")]
    internal class AccessDatabaseTool
    {
        private string mdbFileNameWithPath; //the full path and extension to the .mdb access database file EX: C:/database.mdb
        public AccessDatabaseTool(string mdbFileNameWithPath)
        {
            this.mdbFileNameWithPath = mdbFileNameWithPath;
        }

        /// <summary>
        /// Checks if atleast one row within a table exists
        /// </summary>
        /// <param name="tableName">Name of the Table</param>
        /// <param name="columnName">Name of the Column</param>
        /// <param name="value">Row value to search for in the specified column</param>
        /// <returns></returns>
        public bool RowExists(string tableName, string columnName, string value)
        {
            DataTable output = GetRows(tableName, columnName, value);
            if(output == null) { return false; }
            return output.Rows.Count != 0;
        }
        public bool RowExists(string tableName, KeyValuePair<string, string>[] conditions)
        {
            DataTable output = GetRows(tableName, conditions);
            if (output == null) { return false; }
            return output.Rows.Count != 0;
        }

        /// <summary>
        /// Gets rows where: at the column of conditions.key == conditions.value
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="conditions">array of column,value that must be true in the SQL query</param>
        /// <returns></returns>
        public DataTable GetRows(string tableName, KeyValuePair<string, string>[] conditions)
        {
            var myDataTable = new DataTable();
            using (var conection = new OleDbConnection("Provider = Microsoft.JET.OLEDB.4.0;  Data Source = " + mdbFileNameWithPath))
            {
                conection.Open();
                var query = $"Select * From [{tableName}] Where";
                for (int i = 0; i < conditions.Length; i++)
                {
                    KeyValuePair<string, string> condition = conditions[i];
                    query += $" {condition.Key} = {condition.Value} ";
                    if(i < conditions.Length - 1)
                    { query += " AND "; }
                }
                var adapter = new OleDbDataAdapter(query, conection);
                try { adapter.Fill(myDataTable); }
                catch (Exception ex) { return null; }
                return myDataTable;
            }
        }
        public DataTable GetRows(string tableName, string columnName, string value)
        { return GetRows(tableName, [new KeyValuePair<string, string>(columnName, value)]); }

        /// <summary>
        /// Adds a row into a table and sets one column to a initial value
        /// </summary>
        /// <param name="tableName">Name of the Table</param>
        /// <param name="initialColumn">Column to insert the initialColumnValue</param>
        /// <param name="initialColumnValue">The initial value to insert</param>
        public void AddRow(string tableName, string initialColumn, string initialColumnValue)
        { AddRow(tableName, [new KeyValuePair<string, string>(initialColumn, initialColumnValue)]); }

        public void AddRow(string tableName, KeyValuePair<string, string>[] properties)
        {
            var con = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + mdbFileNameWithPath);
            var cmd = new OleDbCommand();
            cmd.Connection = con;

            string columns = "";
            string values = "";
            for (int i = 0; i < properties.Length; i++)
            {
                KeyValuePair<string, string> property = properties[i];
                columns += $" [{property.Key}] ";
                values += $" '{property.Value}' ";
                if(i != properties.Length - 1)
                {
                    columns += ",";
                    values += ",";
                }
            }
            cmd.CommandText = $"insert into [{tableName}] ({columns})  values ({values});";
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        /// <summary>
        /// Updates a row with a given column and row value
        /// </summary>
        /// <param name="table">Name of the Table</param>
        /// <param name="column">Column to update</param>
        /// <param name="row">Row to update</param>
        /// <param name="properties">An array of KeyValuePairs where the .Key represents the column name and the .Row represents the value.  FOR EXAMPLE: .Key=Name, .Value=Bob</param>
        public void UpdateRow(string table, string column, string row, KeyValuePair<string, string>[] properties)
        {
            var con = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + mdbFileNameWithPath);
            var cmd = new OleDbCommand();
            cmd.Connection = con;

            //build the SQL query of properties in a string
            string propertiesQuery = "";
            for (int i = 0; i < properties.Length; i++)
            {
                KeyValuePair<string, string> property = properties[i];
                propertiesQuery += $"{property.Key} = '{property.Value}'";
                if(i != properties.Length - 1)
                { propertiesQuery += ", "; } //add a comma if there are more properties after this
            }

            cmd.CommandText = $"UPDATE [{table}] SET {propertiesQuery} WHERE {column} = {row};";
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        //modified from : https://stackoverflow.com/questions/8625569/inserting-and-updating-data-to-mdb
    }
}
