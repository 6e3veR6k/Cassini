using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cassini.Model
{
    public class Direction
    {
        public System.Guid Guid { get; set; }
        public string Code { get; set; }
        public string Title { get; set; }

        public string FullName => $"{this.Code} {this.Title}";
    }
}
