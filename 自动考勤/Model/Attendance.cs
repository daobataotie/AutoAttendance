using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAttendance.Model
{
    public class Attendance
    {
        public Attendance(string date, string record, int i)
        {
            this.Date = date;
            
            if (!string.IsNullOrEmpty(record))
            {
                if (i == 1)
                {
                    string[] times = record.Split(' ').Where(R => !string.IsNullOrEmpty(R)).ToArray();
                    if (times.Length > 0)
                    {
                        if (times.Length == 1)
                        {
                            DateTime dt = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + times[0]);
                            DateTime beginDT = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + Common.BeginTime);
                            if (dt > beginDT)
                            {
                                if (dt.Hour - beginDT.Hour > 4)
                                    _afternoon = dt;
                                else
                                    _morning = dt;
                            }
                            else
                            {
                                _morning = dt;
                            }
                        }
                        else if (times.Length > 1)
                        {
                            _morning = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + times[0]);
                            _afternoon = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + times[times.Length - 1]);
                        }
                    }
                    else
                        NoAttendance = true;
                }
                else
                {
                    string[] times = record.Split('\n').Where(R => !string.IsNullOrEmpty(R)).ToArray();
                    if (times.Length > 0)
                    {
                        if (times.Length == 1)
                        {
                            DateTime dt = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + times[0]);
                            DateTime beginDT = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + Common.BeginTime);
                            if (dt > beginDT)
                            {
                                if (dt.Hour - beginDT.Hour > 4)
                                    _afternoon = dt;
                                else
                                    _morning = dt;
                            }
                            else
                            {
                                _morning = dt;
                            }
                        }
                        else if (times.Length > 1)
                        {
                            _morning = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + times[0]);
                            _afternoon = DateTime.Parse(Common.Year + "-" + Common.Month + "-" + date + " " + times[times.Length - 1]);
                        }
                    }
                    else
                        NoAttendance = true;
                }
            }
            else
                NoAttendance = true;
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

        private DateTime? _morning;

        public DateTime? Morning
        {
            get
            {
                return _morning;
            }

            set
            {
                _morning = value;
            }
        }

        private DateTime? _afternoon;

        public DateTime? Afternoon
        {
            get
            {
                return _afternoon;
            }

            set
            {
                _afternoon = value;
            }
        }

        public bool NoAttendance { get; set; } = false;

        private double _normalDay = 0;

        public double NormalDay
        {
            get
            {
                _normalDay = 0;
                if (_morning.HasValue)
                {
                    _normalDay += 0.5;
                }
                if (_afternoon.HasValue)
                {
                    _normalDay += 0.5;
                }
                return _normalDay;
            }
        }

    }
}
