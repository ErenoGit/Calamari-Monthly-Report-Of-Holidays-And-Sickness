using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalamariMonthlyReportOfHolidaysAndSickness
{
    class SingleEmployeeReport
    {
        public long id { get; set; }
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string email { get; set; }

        public AbsencesObject absences { get; set; }
    }
}
