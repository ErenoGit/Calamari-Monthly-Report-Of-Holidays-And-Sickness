using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalamariMonthlyReportOfHolidaysAndSickness
{
    class AbsencesObject
    {
        public string holidays { get; set; }
        public string sickness { get; set; }

        public AbsencesObject()
        {
            holidays = "";
            sickness = "";
        }
    }
}
