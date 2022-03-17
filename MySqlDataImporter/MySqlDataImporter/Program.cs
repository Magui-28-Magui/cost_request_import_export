using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

using MySql.Data.MySqlClient;
using ImportExportDataLibrary;

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

            Import importer = new Import(server, dbuser, dbpassword, databaseName);

            ImportExportDataLibrary.Response response = importer.ImportToDatabase(fileName, sheetName, columns, databaseTableName, fields, types);

            if (response.Success)
            {
                Console.WriteLine("Imported!!!!");
            }
            else
            {
                Console.Error.WriteLine(response.Message);
            }

        }


    }
}
