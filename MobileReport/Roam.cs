using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileReport
{
    internal class Roam
    {
        public object UserNumber { get; set; }
        public object UserName { get; set; }
        public object Description { get; set; }
        public object Amount { get; set; }

        public Roam(object usernumber, object username, object description, object cost)
        {
            UserNumber = usernumber;
            UserName = username;
            Description = description;
            Amount = cost;
        }
    }
}
