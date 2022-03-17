using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExportDataLibrary
{
    public class Export : IConnection
    {
     

            public Export(string server, string user, string pass, string database)
            {
                this.Server = server;
                this.User = user;
                this.Password = pass;
                this.Database = database;
                this.ConnectionString = $"server={server};uid={User};pwd={Password};database={Database}";
            }


        public Response ExportToExcel(string ExcelFileName,
             string ExcelSheetName,
             string[] ExcelColumns,
             string DbTable,
             string[] DbColumns,
             string[] DbTypes, int startRow = 1, int endRow = -1)
        {

            return new Response(false, "Undefined");
        }


        } //end of class

    } //end of namespace
