using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSVtoJSON_Application
{
    class CSVFileFields
    {
        public string firstName { get; set; }
        
        public string Surname { get; set; }

        public string Email { get; set; }

        public string Username { get; set; }

        public string Password { get; set; }

        public string Roles { get; set; }

        public string Groups { get; set; }

        public string organisationUnits { get; set; }

        public string OUCapture { get; set; }
    }
}
