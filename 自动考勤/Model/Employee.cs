using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAttendance.Model
{
    public class Employee
    {
        private string _info;

        [Mapping(ColumnIndex = 2, RowName = "员 工")]
        public string Info
        {
            get
            {
                return _info;
            }

            set
            {
                _info = value;

                if (!string.IsNullOrEmpty(_info))
                {
                    //工号: 1  姓名: 李婷    部门: 整形
                    string[] infos = _info.Split(':');
                    this._id = infos[1].Trim().Split(' ')[0];
                    this._name = infos[2].Trim().Split(' ')[0];
                    this._department = infos[3].Trim().Split(' ')[0];
                }
            }
        }

        private string _id;

        public string Id
        {
            get
            {
                //if (!string.IsNullOrEmpty(_info))
                //{
                //    //工号: 1  姓名: 李婷    部门: 整形
                //    string[] infos = _info.Split(':');
                //    return infos[1].Trim().Split(' ')[0];
                //}
                //else
                //    return "";
                return _id;
            }

            set
            {
                _id = value;
            }
        }

        private string _name;

        public string Name
        {
            get
            {
                //if (!string.IsNullOrEmpty(_info))
                //{
                //    string[] infos = _info.Split(':');
                //    return infos[2].Trim().Split(' ')[0];
                //}
                //else
                //    return "";
                return _name;
            }

            set
            {
                _name = value;
            }
        }

        private string _department;

        public string Department
        {
            get
            {
                //if (!string.IsNullOrEmpty(_info))
                //{
                //    string[] infos = _info.Split(':');
                //    return infos[3].Trim().Split(' ')[0];
                //}
                //else
                //    return "";
                return _department;
            }

            set
            {
                _department = value;
            }
        }

        private string _leader;

        public string Leader
        {
            get
            {
                return _leader;
            }

            set
            {
                _leader = value;
            }
        }

        private string _date;

        public string Date
        {
            get
            {
                return _date;
            }

            set
            {
                _date = value;
            }
        }

        private string _record;

        public string Record
        {
            get
            {
                return _record;
            }

            set
            {
                _record = value;
            }
        }

        List<Attendance> _attendanceList = new List<Attendance>();

        public List<Attendance> AttendanceList
        {
            get
            {
                return _attendanceList;
            }

            set
            {
                _attendanceList = value;
            }
        }


        public double NormalDays
        {
            get
            {
                if (AttendanceList != null && AttendanceList.Count > 0)
                {
                    return AttendanceList.Select(a => a.NormalDay).Sum();
                }
                else
                    return 0;
            }
        }

        public int LateTime
        {
            get
            {
                if (AttendanceList != null && AttendanceList.Count > 0)
                {
                    return AttendanceList.Count(a => a.IsLate);
                }
                else
                    return 0;
            }
        }
    }
}
