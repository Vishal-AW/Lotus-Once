using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Newtonsoft.Json.Linq;

namespace LotusTimesheet.Class
{
    public class EmployeeModel
    {
        public string empno { get; set; }
        public string name { get; set; }
        public string office_email { get; set; }
        public string attendance_date { get; set; }
        public string checkin_time { get; set; }
        public string checkout_time { get; set; }
        public string value { get; set; }
        public string hours { get; set; }
        public string comment { get; set; }
        public string ID { get; set; }

        public int count { get; set; }

    }
}
