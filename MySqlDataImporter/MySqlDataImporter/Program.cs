using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

using MySql.Data.MySqlClient;

namespace MySqlDataImporter
{
    class Program
    {


        static void Main(string[] args)
        {
            //"server=127.0.0.1;uid=root;pwd=12345;database=test"
            //server, user, password, database

            //0. xlsx full filename
            string fileName = args[0];
            //1. sheet name
            string sheetName = args[1];
            //2. columns
            string[] columns = args[2].Split('|');

            Dictionary<string, int> dicColumns = new Dictionary<string, int>();
            foreach (string col in columns)
            {
                dicColumns.Add(col, -1);
            }

            //3. database name
            string databaseName = args[3];

            //4. table
            string databaseTableName = args[4];

            //5. fields
            string[] fields = args[5].Split('|');

            //6. types : int, float, double, string
            string[] types = args[6].Split('|');


            string server = args[7];

            string dbuser = args[8];

            string dbpassword = args[9];

            string connectionString = $"server={server};uid={dbuser};pwd={dbpassword};database={databaseName}";

            //"server=127.0.0.1;uid=root;pwd=12345;database=test"

            if (columns.Length != fields.Length)
            {
                Console.Error.WriteLine("Numbers of columns in worksheet is not equal to numbers of fields in database");
                return;
            }

            //string filename = @"c:\test.xlsx";
            //string sheetname = "FORMATO 1A";
            //string columns = "";

            XSSFWorkbook xssfwb;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheet(sheetName);

            int iRow = 0;
            int iColumn = 0;
            IRow currentRow = sheet.GetRow(iRow);
            int nColumns = currentRow.LastCellNum;
            int nRows = sheet.LastRowNum;

            //Locate column index
            //for(int column = 0; column <= sheet.C ;column++)
            for (int c = 0; c < nColumns; c++)
            {
                ICell cell = currentRow.GetCell(c);

                if (cell.CellType == CellType.String)
                {
                    string cellValue = cell.StringCellValue;

                    if (dicColumns.ContainsKey(cellValue))
                    {
                        //Si se encontro en el dictionario entonces vamos a ponerle numero de la columns
                        dicColumns[cellValue] = c;
                    }
                }
            }


            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();


            MySqlTransaction transaction = connection.BeginTransaction();

            MySqlCommand command = connection.CreateCommand();
            command.Connection = connection;
            command.Transaction = transaction;

            string cSql = "INSERT INTO " + databaseTableName + " ";
            string columnPart = "(";
            string valuesPart = "(";
            //dicColumns.Count()
            for (int f = 0; f < fields.Length; f++)
            {
                string dbColumn = fields[f];

                string colPart = dbColumn;
                string valPart = "@" + dbColumn;

                if (f != fields.Length - 1)
                {
                    colPart += colPart + ", ";
                    valPart += valPart + ", ";
                }

                columnPart += colPart;
                valuesPart += valPart;


            }

            columnPart += ")";
            valuesPart += ")";

            cSql += columnPart + " VALUES " + valuesPart;

            command.CommandText = cSql;

            for (int r = 0; r < nRows; r++)
            {
                //Get all data from each row...
                //Set all parametesr

                //crear todos los parametros
                command.Parameters.Clear();
                for (int f = 0; f < fields.Length; f++)
                {
                    //Parte para crear el parametro
                    MySqlParameter p = new MySqlParameter();
                    p.ParameterName = "@" + fields[f];


                    string nameInExcel = columns[f];
                    int columnIndex = dicColumns[nameInExcel];

                    ICell cell = sheet.GetRow(r).GetCell(columnIndex);
                    p.Value = cell.StringCellValue;

                    string strType = types[f];

                    switch(strType)
                    {
                        case "int":
                            { p.DbType = System.Data.DbType.Int32; }
                            break;
                        case "double":
                            { p.DbType = System.Data.DbType.Double; }
                            break;
                        default:
                            { p.DbType = System.Data.DbType.String; }
                            break;
                    }

                   

                    command.Parameters.Add(p);
                }


                command.ExecuteNonQuery();
            }


            transaction.Commit();

            connection.Close();

        


        }


    }
}
