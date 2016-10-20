using System;
using System.Collections.Generic;
using System.Linq;
using OleDb = System.Data.OleDb;
using Forms = System.Windows.Forms;

namespace StopLight
{
    /*
    Ole Db doesn't support dynamic sql(?),
    but still not sure if there are other tools that can
    */
    internal class OleDbHelper
    {
        //constants for read/write option
        private String _READ = "read";
        private String _WRITE = "write";

        public OleDbHelper()
        {

        }

        internal void ProcessFile(string fileName, ref List<string> indexNames, ref Dictionary<string, int> dictionary)
        {
            //ReadData, GetColumnName, GetWordIndex

            string connectionString = this.PrepareConnectionString(fileName, _READ);

            //'using' acts like a try catch block for reading files
            using (OleDb.OleDbConnection conn = new OleDb.OleDbConnection(connectionString))
            {
                conn.Open();
                this.SaveColumnNames(conn, ref indexNames);
                this.SaveColumnValues(conn, indexNames, ref dictionary);
                conn.Close();
            }

        }
        
        private string PrepareConnectionString(string fileName, string readOrWrite)
        {
            Tuple<string, string>[] connectionParameters = new Tuple<string, string>[3];

            connectionParameters[0] = new Tuple<string, string>("Provider", "Microsoft.ACE.OLEDB.12.0");
            connectionParameters[1] = new Tuple<string, string>("Data Source", fileName);
            //HDR = YES : whether the first row should be header names
            string header = ";HDR=YES";
            //IMEX = 1 : parse as strings
            string parseAsString = ";IMEX=";
            if (readOrWrite.Equals(_READ))//read only
                parseAsString += "1";
            else if (readOrWrite.Equals(_WRITE))
                parseAsString += "0";
            /*
            The header 'excelType' given below is for Microsoft Excel 2010 and 2007 only
            Change the code so that if you are running another version of Excel (for example 2003)
            We will change the connection parameters
            Check this link for more information on the headers
            https://www.connectionstrings.com/net-framework-data-provider-for-ole-db/
            How to detect execl versions is up to you.
            */
            string excelType = "Excel 12.0";
            /*
            CODE HERE
            */
            connectionParameters[2] = new Tuple<string, string>("Extended Properties", "\"" + excelType + header + parseAsString + "\"");

            string connectionString = "";
            foreach (Tuple<string, string> parameter in connectionParameters)
                connectionString += parameter.Item1 + "=" + parameter.Item2 + ";";

            return connectionString;
        }

        private void SaveColumnNames(OleDb.OleDbConnection conn, ref List<string> indexNames)
        {
            //Has to be called BEFORE SaveColumnValues()
            System.Data.DataTable columns = conn.GetSchema("Columns");
            //by default column names are sorted in alphabetical order
            //(this is a stupid feature)
            //so sort back to order in Excel
            System.Data.DataRow[] rows = columns.Select("", "ORDINAL_POSITION");

            foreach (System.Data.DataRow row in rows)
            {
                string columnName = (string)row["COLUMN_NAME"];
                indexNames.Add(columnName);
            }
        }

        private void SaveColumnValues(OleDb.OleDbConnection conn, List<string> indexNames, ref Dictionary<string, int> dictionary)
        {
            //Call ReadColumnNames first!
            int indexNumber = 0;
            //just making sure column names are populated
            if (indexNames.Count() == 0)
                return;

            foreach (string index in indexNames)
            {
                //no need to prepare this command
                OleDb.OleDbCommand selectAll = new OleDb.OleDbCommand("SELECT `" + index + "` FROM `Sheet1$`", conn);
                try {
                    OleDb.OleDbDataReader reader = selectAll.ExecuteReader();
                    while (reader.Read())
                    {
                        //potentially throws invalid cast exception (null)
                        if (!reader.IsDBNull(0))
                        {
                            string word = reader.GetString(0);
                            //clean word
                            string cleanWord = Globals.ThisAddIn.HighlightManager.CleanWord(word);
                            //save word
                            if (Globals.ThisAddIn.HighlightManager.CanAddToDictionary(cleanWord))
                                dictionary.Add(cleanWord, indexNumber);
                        }
                    }
                    reader.Close();
                } catch (OleDb.OleDbException)
                {
                    Forms.MessageBox.Show(Strings.error, Strings.errorCaption);
                }
                indexNumber++;
            }
        }

        internal void AddWordToExcel(string fileName, string word, string columnName)
        {
            //this should be handled in the form to add the word,
            //but just in case
            if (!Globals.ThisAddIn.HighlightManager.CanAddToDictionary(word))
            {
                return;
            }
            string connectionString = this.PrepareConnectionString(fileName, _WRITE);
            using (OleDb.OleDbConnection conn = new OleDb.OleDbConnection(connectionString))
            {
                conn.Open();
                OleDb.OleDbCommand insert;
                if (ShouldInsertIntoNewRow(conn, columnName))
                    insert = PrepareInsertNewRowCommand(conn, columnName);
                else
                    insert = PrepareInsertExistingRowCommand(conn, columnName);
                insert.Parameters[0].Value = word;
                try
                {
                    insert.ExecuteNonQuery();
                } catch (OleDb.OleDbException)
                {
                    Forms.MessageBox.Show(Strings.error, Strings.errorCaption);
                }

                conn.Close();
            }
        }

        // column1 column2 column3
        // ****    ***     *****
        //  **              ***
        // If inserting into column 2, we are inserting into an existing row
        // if inserting into column 1 or 3, we are creating a new row
        private bool ShouldInsertIntoNewRow(OleDb.OleDbConnection conn, string columnName)
        {
            OleDb.OleDbCommand command = new OleDb.OleDbCommand();
            command.Connection = conn;
            command.CommandText = "SELECT TOP 1 * FROM `Sheet1$` WHERE `" + columnName + "` IS NULL";
            int ct = 0;
            try
            {
                OleDb.OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                    ct++;

                reader.Close();
                
            } catch (OleDb.OleDbException)
            {
                Forms.MessageBox.Show(Strings.error, Strings.errorCaption);
            }

            if (ct > 0)
                return false;
            else
                return true;
        }

        private OleDb.OleDbCommand PrepareInsertExistingRowCommand(OleDb.OleDbConnection conn, string columnName)
        {
            OleDb.OleDbCommand command = new OleDb.OleDbCommand();
            command.Connection = conn;
            command.CommandText = "UPDATE (SELECT TOP 1 * FROM `Sheet1$` WHERE `" + columnName + "` IS NULL)"
                + "SET `" + columnName + "` = ?";
            command.Parameters.Add(columnName, OleDb.OleDbType.VarWChar, 50);
            return command;
        }

        private OleDb.OleDbCommand PrepareInsertNewRowCommand(OleDb.OleDbConnection conn, string columnName)
        {
            OleDb.OleDbCommand command = new OleDb.OleDbCommand();
            command.Connection = conn;
            command.CommandText = "INSERT INTO `Sheet1$` (" + columnName + ") VALUES (?)";
            command.Parameters.Add(columnName, OleDb.OleDbType.VarWChar, 50);
            return command;
        }
    }
}
