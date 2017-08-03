using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutoAttendance
{
    public class ExcelOperation
    {
        public List<Model.Employee> GetExcelDate_Read1(string fileName)
        {
            ExcelHelper excelHelper = null;
            Model.Employee entity = new Model.Employee();
            List<Model.Employee> list = new List<Model.Employee>();
            try
            {
                excelHelper = new ExcelHelper(fileName);

                object[,] data = (object[,])excelHelper.GetData(Common.SheetName_Read1, Common.EndColumn_Read1);
                List<Header> headerList = new List<Header>();

                PropertyInfo[] properties = entity.GetType().GetProperties();

                //填充Header
                for (int i = 0; i < properties.Length; i++)
                {
                    PropertyInfo item = properties[i];
                    if (properties[i].CustomAttributes != null && properties[i].CustomAttributes.Count<CustomAttributeData>() != 0)
                    {
                        CustomAttributeData ca = item.CustomAttributes.First<CustomAttributeData>();
                        Header header = new Header();
                        //header.PropertyInfo = item;
                        foreach (CustomAttributeNamedArgument na in ca.NamedArguments)
                        {
                            if (na.MemberName == "ColumnIndex")
                            {
                                header.ColumnIndex = int.Parse(na.TypedValue.Value.ToString());
                            }
                            if (na.MemberName == "RowName")
                            {
                                header.RowName = na.TypedValue.Value.ToString();
                            }
                        }
                        if (!headerList.Any(h => h.ColumnIndex == header.ColumnIndex && h.RowName == header.RowName))
                        {
                            headerList.Add(header);
                        }
                    }
                }
                bool needAddHeader = true;
                for (int i = int.Parse(Common.StartRow_Read1.ToString()); i <= data.GetLength(0); i++)
                {
                    for (int j = 1; j <= data.GetLength(1); j++)
                    {
                        if (data[i, j] != null)
                        {
                            //var header = headerList.Where(h => data[i, j].ToString().Contains(h.ColumnName)).ToList();
                            //if (header != null && header.Count > 0)
                            //{
                            //    header.ForEach(h => h.ColumnIndex = j);
                            //}
                            if (data[i, j].ToString() == "员 工")
                            {
                                entity = new Model.Employee();
                                list.Add(entity);

                                var headerUpdate = headerList.Where(h => data[i, j].ToString().Contains(h.RowName) && h.ColumnIndex != 0).ToList();
                                if (headerUpdate != null && headerUpdate.Count > 0)
                                {
                                    headerUpdate.ForEach(h =>
                                    {
                                        entity.Info = data[i, h.ColumnIndex]?.ToString().Trim();
                                        //h.PropertyInfo.SetValue(entity, data[i, h.ColumnIndex]?.ToString().Trim());
                                    });
                                }

                                break;
                            }
                            //填充Header
                            else if (data[i, j].ToString() == "日 期" && needAddHeader)
                            {
                                for (int tem = 2; tem < data.GetLength(1); tem++)
                                {
                                    if (data[i, tem] != null && int.Parse(data[i, tem].ToString()) <= Common.TotalDays)
                                        headerList.Add(new Header() { RowName = "日 期", ColumnName = data[i, tem].ToString().Trim(), ColumnIndex = tem });
                                }
                                if (headerList.Count(h => h.RowName == "日 期") >= 28)
                                    needAddHeader = false;

                                break;
                            }
                            else if (data[i, j].ToString() == "考 勤")
                            {
                                var headerDate = headerList.Where(h => h.RowName == "日 期").ToList();
                                if (headerDate != null && headerDate.Count > 0)
                                {
                                    foreach (var item in headerDate)
                                    {
                                        if (data[i - 1, item.ColumnIndex] != null && data[i - 1, item.ColumnIndex].ToString() == item.ColumnName)
                                        {
                                            entity.Date = item.ColumnName;
                                            entity.Record = data[i, item.ColumnIndex]?.ToString().Trim();

                                            entity.AttendanceList.Add(new Model.Attendance(entity.Date, entity.Record, 1));
                                        }
                                    }
                                }

                                break;
                            }
                        }
                        //else if (j == 1 && data[i, j] == null)
                        //{
                        //    if (!string.IsNullOrEmpty(entity.Id))
                        //    {
                        //        list.Add(entity);
                        //        entity = new Model.Employee();
                        //    }

                        //    break;
                        //}
                    }
                }
                return list;

            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (excelHelper != null)
                {
                    excelHelper.Close();
                }
            }
        }

        public List<Model.Employee> GetExcelDate_Read2(string fileName)
        {
            ExcelHelper excelHelper = null;
            Model.Employee entity = new Model.Employee();
            List<Model.Employee> list = new List<Model.Employee>();
            try
            {
                excelHelper = new ExcelHelper(fileName);
                object[,] data = (object[,])excelHelper.GetData(Common.SheetName_Read2, Common.EndColumn_Read2);

                List<Header> headerList = new List<Header>();

                for (int i = int.Parse(Common.StartRow_Read2.ToString()); i <= data.GetLength(0); i++)
                {
                    if (i == int.Parse(Common.StartRow_Read2.ToString()))
                    {
                        for (int j = 1; j <= data.GetLength(1); j++)
                        {
                            if (data[i, j] != null && !headerList.Any(H => H.ColumnName == data[i, j].ToString() && H.ColumnIndex == j))
                                headerList.Add(new Header { ColumnName = data[i, j].ToString(), ColumnIndex = j });
                        }
                    }
                    else
                    {
                        for (int j = 1; j <= data.GetLength(1); j++)
                        {
                            if (data[i, j] != null && headerList.Any(H => H.ColumnName == data[i, j].ToString() && H.ColumnIndex == j))
                                break;
                            else if (data[i, j] != null && data[i, j].ToString().Contains("工 号："))
                            {
                                entity = new Model.Employee();
                                entity.Id = data[i, 3].ToString().Trim();
                                entity.Name = data[i, 11].ToString().Trim();
                                entity.Department = data[i, 21].ToString().Trim();

                                list.Add(entity);
                                break;
                            }
                            else if (j <= Common.TotalDays)
                            {
                                Model.Attendance attendance = new Model.Attendance(j.ToString(), data[i, j]?.ToString(), 2);
                                entity.AttendanceList.Add(attendance);
                            }
                        }
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (excelHelper != null)
                {
                    excelHelper.Close();
                }
            }
        }

        public void WriteExcel(string fileName, Dictionary<string, string> dic, List<Model.Employee> empList)
        {
            ExcelHelper excelHelper = null;
            try
            {
                excelHelper = new ExcelHelper(fileName);
                foreach (var item in excelHelper.sheets)
                {
                    dynamic range = excelHelper.GetRange(item, Common.EndColumn_Write);
                    object[,] data = range.Value2;

                    string employeeName = string.Empty;
                    Model.Employee emp = null;
                    for (int i = 1; i <= data.GetLength(0); i++)
                    {
                        string morningOrAfternoon = string.Empty;
                        for (int j = 1; j <= data.GetLength(1); j++)
                        {
                            //填充日期
                            if (i == 1 && j >= int.Parse(Common.StartColumn_WriteWeek) && j < int.Parse(Common.StartColumn_WriteWeek) + Common.TotalDays)
                            {
                                string value = dic[(j - int.Parse(Common.StartColumn_WriteWeek) + 1).ToString()];
                                switch (value)
                                {
                                    case "六":
                                        excelHelper.SetCellValueWithColor(item, i, j, value, Common.Color.Green);

                                        //修改星期颜色时，把对应的日期颜色也修改
                                        excelHelper.SetCellColor(item, i + 1, j, Common.Color.Green);
                                        break;
                                    case "日":
                                        excelHelper.SetCellValueWithColor(item, i, j, value, Common.Color.Red);
                                        excelHelper.SetCellColor(item, i + 1, j, Common.Color.Red);
                                        break;
                                    default:
                                        excelHelper.SetCellValue(item, i, j, value);
                                        break;
                                }
                            }
                            //填充考勤
                            else if (i >= int.Parse(Common.StartRow_Write))
                            {
                                if (j == 1 && data[i, j] != null)
                                {
                                    employeeName = data[i, j].ToString().Trim();
                                    emp = empList.FirstOrDefault(e => e.Name == employeeName);

                                    if (emp == null)
                                        break;

                                    continue;
                                }

                                if (emp != null)
                                {
                                    if (j == 2 && data[i, j] != null)
                                    {
                                        morningOrAfternoon = data[i, j].ToString().Trim();

                                        continue;
                                    }
                                    else if (j >= 3 && j < 3 + Common.TotalDays)
                                    {
                                        Model.Attendance attendance = emp.AttendanceList.Single(A => A.Date == (j - 2).ToString());
                                        if (!attendance.NoAttendance)
                                        {
                                            DateTime? dt = null;

                                            if (morningOrAfternoon == "上午")
                                                dt = emp.AttendanceList.Single(A => A.Date == (j - 2).ToString()).Morning;
                                            else if (morningOrAfternoon == "下午")
                                                dt = emp.AttendanceList.Single(A => A.Date == (j - 2).ToString()).Afternoon;
                                            else
                                                break;

                                            if (dt.HasValue)
                                            {
                                                if (attendance.IsLate && morningOrAfternoon == "上午")
                                                    excelHelper.SetCellValue(item, i, j, "√", Common.Color.Red);
                                                else
                                                    excelHelper.SetCellValue(item, i, j, "√");
                                            }
                                            //else
                                            //    excelHelper.SetCellValue(item, i, j, "√", Common.Color.Red);
                                        }
                                    }
                                    else if (j == 3 + 31)
                                    {
                                        excelHelper.SetCellValue(item, i, j, emp.NormalDays == 0 ? "" : emp.NormalDays.ToString());
                                    }
                                    else if (j == 3 + 36)
                                        excelHelper.SetCellValue(item, i, j, emp.LateTime == 0 ? "" : emp.LateTime.ToString());
                                }
                                else
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                excelHelper.Save();
                if (excelHelper != null)
                {
                    excelHelper.Close();
                }
            }
        }

    }

    internal class Header
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

        //public PropertyInfo PropertyInfo
        //{
        //    get;
        //    set;
        //}

    }
}
