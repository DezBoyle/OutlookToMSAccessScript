using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Net;

namespace OutlookToMSAccessScript
{
    internal class AccessDatabaseTool
    {
        private string fileNameWithPath;
        public AccessDatabaseTool(string fileNameWithPath)
        {
            this.fileNameWithPath = fileNameWithPath;
        }
        public void Test()
        {
            var myDataTable = new DataTable();
            using (var conection = new OleDbConnection("Provider = Microsoft.JET.OLEDB.4.0;  Data Source = " + fileNameWithPath))
            {
                conection.Open();
                var query = "Select CompanyName From CompanyName";
                var adapter = new OleDbDataAdapter(query, conection);
                adapter.Fill(myDataTable);
                Console.WriteLine(myDataTable.Rows[0][0].ToString());
                Console.ReadKey();
            }
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
            var myDataTable = new DataTable();
            using (var conection = new OleDbConnection("Provider = Microsoft.JET.OLEDB.4.0;  Data Source = " + fileNameWithPath))
            {
                conection.Open();
                var query = $"Select * From {tableName} Where {columnName} = '{value}'";
                var adapter = new OleDbDataAdapter(query, conection);
                adapter.Fill(myDataTable);
                return myDataTable.Rows.Count != 0;
            }
        }

        /// <summary>
        /// Adds a row into a table and sets one column to a initial value
        /// </summary>
        /// <param name="tableName">Name of the Table</param>
        /// <param name="initialColumn">Column to insert the initialColumnValue</param>
        /// <param name="initialColumnValue">The initial value to insert</param>
        public void AddRow(string tableName, string initialColumn, string initialColumnValue)
        {
            var con = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + fileNameWithPath);
            var cmd = new OleDbCommand();
            cmd.Connection = con;
            cmd.CommandText = $"insert into {tableName} ({initialColumn})  values ('{initialColumnValue}');";
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        public void UpdateRow(string table, string column, string row, KeyValuePair<string, string>[] properties)
        {
            var con = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + fileNameWithPath);
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

            cmd.CommandText = $"UPDATE CompanyName SET {propertiesQuery} WHERE {column} = '{row}';";
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        //modified from : https://stackoverflow.com/questions/8625569/inserting-and-updating-data-to-mdb
    }
}
