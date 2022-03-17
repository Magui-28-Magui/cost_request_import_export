using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

using MySql.Data.MySqlClient;
using System.IO;
using System.Diagnostics;

namespace ImportExportDataLibrary
{

    //Import to mysql
    public class Import : IConnection
    {


        public Import(string server, string user, string pass, string database)
        {
            this.Server = server;
            this.User = user;
            this.Password = pass;
            this.Database = database;
            this.ConnectionString = $"server={server};uid={User};pwd={Password};database={Database}";
        }


        //Dictionary to save columns and his index
        Dictionary<string, int> dicExcelColumns = new Dictionary<string, int>();

        public ImportExportDataLibrary.Response ImportToDatabase(string ExcelFileName,
            string ExcelSheetName,
            string[] ExcelColumns,

            string DbTable,
            string[] DbColumns,
            string[] DbTypes, int startRow = 1, int endRow = -1)
        {

            if (ExcelColumns.Length != DbColumns.Length)
            {
                return new Response(false, "Excel Columns and MySql table Columns are not equals in number.");
            }

            FileStream file = new FileStream(ExcelFileName, FileMode.Open, FileAccess.Read);
            XSSFWorkbook xssfwb = new XSSFWorkbook(file);
            ISheet sheet = xssfwb.GetSheet(ExcelSheetName);

            //set row in first to read columns
            int iRow = 0;
            IRow currentRow = sheet.GetRow(iRow);

            int nColumns = currentRow.LastCellNum;
            int nRows = sheet.LastRowNum;

            FillExcelColumnsDictionary(sheet, ExcelColumns, nColumns);

            MySqlConnection connection = new MySqlConnection(this.ConnectionString);
            connection.Open();

            MySqlTransaction transaction = connection.BeginTransaction();

            MySqlCommand command = connection.CreateCommand();
            command.Connection = connection;
            command.Transaction = transaction;

            command.CommandText = getInsertCsql(DbTable, DbColumns);

            if (endRow == -1)
                endRow = nRows - 1;

            var stopWatch = new Stopwatch();
            stopWatch.Start();


            String debugStr = "";
            try
            {
                for (int r = startRow; r <= endRow; r++)
                {
                    debugStr = "Executing row: " + (r + 1) + ", [";

                    //crear todos los parametros
                    command.Parameters.Clear();

                    for (int f = 0; f < DbColumns.Length; f++)
                    {           
                        //Parte para crear el parametro
                        MySqlParameter p = new MySqlParameter();
                        p.ParameterName = "@" + DbColumns[f];

                        string nameInExcel = ExcelColumns[f];

                        int columnIndex = dicExcelColumns[nameInExcel];

                        ICell cell = sheet.GetRow(r).GetCell(columnIndex);

                        if (cell == null)
                        {
                            p.Value = System.DBNull.Value;
                            debugStr += "null, ";
                        }
                        else
                        {
                            if (cell.CellType == CellType.Numeric)
                                p.Value = cell.NumericCellValue;
                            else if (cell.CellType == CellType.String)
                                p.Value = cell.StringCellValue;

                            if(p.Value == null)
                                debugStr +=  "null, ";
                            else
                            debugStr += p.Value.ToString() + ", ";
                        }

                        

                        string strType = DbTypes[f];

                        switch (strType)
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

                    debugStr += "]";

                    command.ExecuteNonQuery();
                }

            }
            catch (Exception ex)
            {
                connection.Close();
                return new Response(false, "Error Ocurred (" + debugStr + "): " + ex.Message );
            }


            transaction.Commit();
            connection.Close();

            stopWatch.Stop();
            

            return new Response(true, "Success import! in " + stopWatch.ElapsedMilliseconds );
        }


        private string getInsertCsql(string table, string[] fields)
        {
            string cSql = "INSERT INTO " + table + " ";
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
                    colPart += ", ";
                    valPart += ", ";
                }

                columnPart += colPart;
                valuesPart += valPart;
            }

            columnPart += ")";
            valuesPart += ")";

            cSql += columnPart + " VALUES " + valuesPart;

            return cSql;
        }


        private void FillExcelColumnsDictionary(ISheet sheet, string[] ExcelColumns, int nColumns)
        {
            IRow currentRow = sheet.GetRow(0);

            //borrar el dictionary para mantenerlo vacio
            dicExcelColumns.Clear();
            foreach (string col in ExcelColumns)
            {
                dicExcelColumns.Add(col, -1);
            }



            //Locate column index
            for (int c = 0; c < nColumns; c++)
            {
                ICell cell = currentRow.GetCell(c);

                if (cell.CellType == CellType.String)
                {
                    string cellValue = cell.StringCellValue;

                    if (dicExcelColumns.ContainsKey(cellValue))
                    {
                        //Si se encontro en el dictionario entonces vamos a ponerle numero de la columns
                        dicExcelColumns[cellValue] = c;
                    }
                }
            }
        }

       
        

    }
}
