using System;
using System.Data;
using System.Data.OleDb;


namespace MobileReport
{
    internal class Connect
    {
        private const string data = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "Data Source=\"{0}\";" +
                                    "Mode=ReadWrite|Share Deny None;" +
                                    "Extended Properties='Excel 12.0; HDR={1}; IMEX={2}';" +
                                    "Persist Security Info=False";

        public DataSet OpenExcelSheet(string FileName, string SheetName)
        {
            string connStr = null;
            string sql1 = "SELECT * FROM [" + SheetName + "$]";
            DataSet dataSet = new DataSet();
            OleDbConnection oleDbConnection = null;

            try
            {
                connStr = string.Format(data, FileName, "YES", "1");
                oleDbConnection = new OleDbConnection(connStr);

                oleDbConnection.Open();
                OleDbCommand cmd = new OleDbCommand(sql1, oleDbConnection);
                OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);
                adpt.Fill(dataSet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oleDbConnection != null) oleDbConnection.Close();
            }
            return dataSet;
        }

        public String[] ExcelSheetNames(string excelFile)
        {
            string connStr = null;
            DataTable schema = new DataTable();
            connStr = string.Format(data, excelFile, "YES", "1");
            OleDbConnection oleDbConnection = new OleDbConnection(connStr);
            oleDbConnection.Open();
            schema = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string[] sheetNames = new string[schema.Rows.Count];
            for (int i = 0; i < schema.Rows.Count; i++)
            {
                sheetNames[i] = schema.Rows[i]["TABLE_NAME"].ToString().Trim('\'').Replace("$", "");
            }
            oleDbConnection.Close();
            return sheetNames;
        }
    }
}
