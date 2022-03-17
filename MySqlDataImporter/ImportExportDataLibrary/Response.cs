using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExportDataLibrary
{
    public class Response
    {
        public Response(bool success, string mess)
        {
            this.Success = success;
            this.Message = mess;
        }

        public bool Success { get; set; }
        public string Message { get; set; }

    }
}
