using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AutoAttendance
{
    public class ExcelHelper
    {
        private dynamic xlApp = null;
        private dynamic workbook = null;
        private int processId = 0;
        public dynamic sheets = null;

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        public ExcelHelper(string fileName)
        {
            Type t = Type.GetTypeFromProgID("EXCEL.Application");
            if (t == null)
            {
                throw new Exception("本机沒有安装Office-Excel");
            }
            this.xlApp = Activator.CreateInstance(t);

            GetWindowThreadProcessId(((IntPtr)this.xlApp.Hwnd), out this.processId);

            this.xlApp.Workbooks.Open(fileName);
            this.workbook = this.xlApp.ActiveWorkbook;
            sheets = this.xlApp.Sheets;
        }

        public void Close()
        {
            this.workbook.Close();
            if (this.processId != 0)
            {
                Process p = Process.GetProcessById(this.processId);
                p.Kill();
            }
        }

        public void Save()
        {
            this.workbook.Save();
        }

        public int GetColumnIndex(string endColumn)
        {
            string columnRef = "ABCDEFGHIGKLMNOPQRSTUVWXYZ";
            return columnRef.IndexOf(endColumn) + 1;
        }

        public object GetData(string endColumn)
        {
            dynamic sheet = this.xlApp.ActiveSheet;

            int rowCount = sheet.UsedRange.Rows.Count;
            if (rowCount == 0)
            {
                throw new Exception("Excel数据错误");
            }
            dynamic range = sheet.Range["A1", string.Format("{0}{1}", endColumn, rowCount)];

            return range.Value2;
        }

        public object GetData(string sheetName, string endColumn)
        {
            dynamic sheet = null;
            try
            {
                sheet = this.xlApp.Sheets[sheetName];
            }
            catch (Exception ex)
            {
                throw new Exception($"Sheet名\"{sheetName}\"无效");
            }

            int rowCount = sheet.UsedRange.Rows.Count;
            if (rowCount == 0)
            {
                throw new Exception("Excel数据错误");
            }
            dynamic range = sheet.Range["A1", string.Format("{0}{1}", endColumn, rowCount)];

            return range.Value2;
        }

        public object GetRange(dynamic sheet, string endColumn)
        {

            int rowCount = sheet.UsedRange.Rows.Count;
            if (rowCount == 0)
            {
                throw new Exception("Excel数据错误");
            }
            dynamic range = sheet.Range["A1", string.Format("{0}{1}", endColumn, rowCount)];

            return range;
        }

        public void SetRangeFormat(string startIndex, string endIndex, object format)
        {
            dynamic sheet = this.xlApp.ActiveSheet;
            dynamic range = sheet.Range[startIndex, endIndex];

            range.NumberFormatLocal = format;
        }

        //public void SetCellValue(int rowIndex, int colIndex, object value)
        //{
        //    dynamic sheet = this.xlApp.ActiveSheet;
        //    dynamic cell = sheet.Cells[rowIndex, colIndex];

        //    cell.Value2 = value;
        //}

        public void SetCellValue(dynamic sheet, int rowIndex, int colIndex, object value)
        {
            dynamic cell = sheet.Cells[rowIndex, colIndex];

            cell.Value2 = value;
        }

        public void SetCellValue(dynamic sheet, int rowIndex, int colIndex, object value,Common.Color color)
        {
            dynamic cell = sheet.Cells[rowIndex, colIndex];
            cell.Font.Color = color;

            cell.Value2 = value;
        }

        public void SetCellValueWithColor(dynamic sheet, int rowIndex, int colIndex, object value, Common.Color color)
        {
            dynamic cell = sheet.Cells[rowIndex, colIndex];
            cell.Interior.Color = color;

            cell.Value2 = value;
        }

        public void SetCellColor(dynamic sheet, int rowIndex, int colIndex, Common.Color color)
        {
            dynamic cell = sheet.Cells[rowIndex, colIndex];
            cell.Interior.Color = color;
        }

        public List<string> SheetNames
        {
            get
            {
                if (this.sheets != null)
                {
                    List<string> list = new List<string>();
                    foreach (dynamic item in this.sheets)
                    {
                        list.Add(item.Name);
                    }
                    return list;
                }
                return null;
            }
        }
    }
}
