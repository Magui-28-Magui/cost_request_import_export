using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExportDataLibrary
{
    public class IConnection
    {


        //Database Properties
        public string Server { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
        public string Database { get; set; }
        public string ConnectionString { get; set; }

    }
}
