using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAttendance
{
    public static class Common
    {
        public static string StartRow_Read1 { get; set; }

        public static string StartRow_Read2 { get; set; }

        public static string EndColumn_Read1 { get; set; }

        public static string EndColumn_Read2 { get; set; }

        public static string SheetName_Read1 { get; set; }

        public static string SheetName_Read2 { get; set; }

        public static string BeginTime { get; set; }

        public static string EndTime { get; set; }

        public static string Year { get; set; }

        public static string Month { get; set; }

        public static int TotalDays
        {
            get
            {
                if (!string.IsNullOrEmpty(Year) && !string.IsNullOrEmpty(Month))
                {
                    return DateTime.DaysInMonth(int.Parse(Year), int.Parse(Month));
                }
                else
                    return 0;
            }
        }

        public static Dictionary<string, string> WeekMapping
        {
            get
            {
                if (TotalDays != 0)
                {
                    Dictionary<string, string> dic = new Dictionary<string, string>();

                    for (int i = 1; i <= TotalDays; i++)
                    {
                        DateTime dt = DateTime.Parse(Year + "-" + Month + "-" + i.ToString());
                        string weekName = GetChineseWeekName(dt.DayOfWeek.ToString());

                        dic.Add(i.ToString(), weekName);
                    }

                    return dic;
                }
                return null;
            }
        }

        private static string GetChineseWeekName(string englishName)
        {
            string str = string.Empty;
            switch (englishName)
            {
                case "Monday":
                    str = "一";
                    break;
                case "Tuesday":
                    str = "二";
                    break;
                case "Wednesday":
                    str = "三";
                    break;
                case "Thursday":
                    str = "四";
                    break;
                case "Friday":
                    str = "五";
                    break;
                case "Saturday":
                    str = "六";
                    break;
                case "Sunday":
                    str = "日";
                    break;
            }
            return str;
        }

        public static string TemplateFileName { get; set; }

        public static string StartRow_WriteWeek { get; set; }

        public static string StartColumn_WriteWeek { get; set; }

        public static string EndColumn_Write { get; set; }

        public static string StartRow_Write { get; set; }

        public enum Color
        {
            Red = 255,
            Green = 3407718,
            Yellow = 65535
        }
    }
}
