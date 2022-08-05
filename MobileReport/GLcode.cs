using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileReport
{
    internal class GLcode
    {
        public object UserNumber { get; set; }
        public object GLCODE { get; set; }
        public object Division { get; set; }
        public object Position { get; set; }
        public object UserName { get; set; }

        public GLcode(object usernumber, object glnumber, object division, object position, object username)
        {
            UserNumber = usernumber;
            GLCODE = glnumber;
            Division = division;
            Position = position;
            UserName = username;
        }
    }
}
