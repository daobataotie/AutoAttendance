using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAttendance.Model
{
    public class MappingAttribute : Attribute
    {
        public string ColumnName
        {
            get;
            set;
        }

        public int ColumnIndex
        {
            get;
            set;
        }

        public string RowName
        {
            get;
            set;
        }

    }
}
