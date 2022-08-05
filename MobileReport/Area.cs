using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileReport
{
    internal class Area
    {
        public object AreaCode { get; set; }
        public object Province { get; set; }
        public object Tax { get; set; }

        public Area(object areacode, object province, object tax)
        {
            AreaCode = areacode;
            Province = province;
            Tax = tax;
        }
    }
}
