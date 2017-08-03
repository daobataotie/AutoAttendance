using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.IO;

namespace AutoAttendance
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;

            InitializeData();
        }

        private void InitializeData()
        {
            this.txt_ExcelPath1.ReadOnly = true;

            Common.StartRow_Read1 = ConfigurationManager.AppSettings["StartRow_Read1"];
            Common.EndColumn_Read1 = ConfigurationManager.AppSettings["EndColumn_Read1"];
            Common.StartRow_Read2 = ConfigurationManager.AppSettings["StartRow_Read2"];
            Common.EndColumn_Read2 = ConfigurationManager.AppSettings["EndColumn_Read2"];

            Common.BeginTime = ConfigurationManager.AppSettings["BeginTime"];
            Common.EndTime = ConfigurationManager.AppSettings["EndTime"];

            this.txt_SheetName1.Text = Common.SheetName_Read1 = ConfigurationManager.AppSettings["SheetName_Read1"];
            this.txt_SheetName2.Text = Common.SheetName_Read2 = ConfigurationManager.AppSettings["SheetName_Read2"];

            Common.StartRow_WriteWeek = ConfigurationManager.AppSettings["StartRow_WriteWeek"];
            Common.StartColumn_WriteWeek = ConfigurationManager.AppSettings["StartColumn_WriteWeek"];
            Common.EndColumn_Write = ConfigurationManager.AppSettings["EndColumn_Write"];
            Common.StartRow_Write = ConfigurationManager.AppSettings["StartRow_Write"];

            Common.TemplateFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", ConfigurationManager.AppSettings["Template"]);

            this.cob_BeginTime.SelectedIndex = this.cob_BeginTime.Items.IndexOf(Common.BeginTime);
            this.cob_EndTime.SelectedIndex = this.cob_EndTime.Items.IndexOf(Common.EndTime);

            DateTime dt = DateTime.Now;
            for (int i = 0; i <= 5; i++)
            {
                this.cob_Month.Items.Add(dt.AddMonths(-i).ToString("yyyy-MM"));
            }
            this.cob_Month.SelectedIndex = this.cob_Month.Items.IndexOf((DateTime.Now.AddMonths(-1)).ToString("yyyy-MM"));
        }

        private void btn_InportExcel1_Click(object sender, EventArgs e)
        {
            this.txt_ExcelPath1.Text = "";

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件|*.xls;*.xlsx";
            ofd.Multiselect = true;
            ofd.Title = "请选择数据源";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in ofd.FileNames)
                {
                    this.txt_ExcelPath1.Text += item + ",";
                }
                this.txt_ExcelPath1.Text = this.txt_ExcelPath1.Text.TrimEnd(',');
                this.txt_ExcelPath1.SelectionStart = this.txt_ExcelPath1.Text.Length;
            }
        }

        private void btn_InportExcel2_Click(object sender, EventArgs e)
        {
            this.txt_ExcelPath2.Text = "";

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件|*.xls;*.xlsx";
            ofd.Multiselect = false;
            ofd.Title = "请选择数据源";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.txt_ExcelPath2.Text = ofd.FileName;
                this.txt_ExcelPath2.SelectionStart = this.txt_ExcelPath2.Text.Length;
            }
        }

        private void btn_SaveSetting_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.txt_SheetName1.Text) || string.IsNullOrEmpty(this.txt_SheetName2.Text))
            {
                MessageBox.Show("Sheet名不能为空！", "提示", MessageBoxButtons.OK);
                return;
            }
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["BeginTime"].Value = this.cob_BeginTime.Text;
            config.AppSettings.Settings["EndTime"].Value = this.cob_EndTime.Text;
            config.AppSettings.Settings["SheetName_Read1"].Value = this.txt_SheetName1.Text;
            config.AppSettings.Settings["SheetName_Read2"].Value = this.txt_SheetName2.Text;

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            MessageBox.Show("保存成功！", "提示", MessageBoxButtons.OK);
        }

        private void btn_ExportReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txt_ExcelPath1.Text) && string.IsNullOrEmpty(this.txt_ExcelPath2.Text))
                {
                    MessageBox.Show("请先导入数据源！", "提示", MessageBoxButtons.OK);
                    return;
                }

                Common.BeginTime = this.cob_BeginTime.Text;
                Common.EndTime = cob_EndTime.Text;
                Common.Year = this.cob_Month.Text.Split('-')[0];
                Common.Month = this.cob_Month.Text.Split('-')[1];

                ExcelOperation eo = new ExcelOperation();

                List<Model.Employee> list10 = new List<Model.Employee>();
                List<Model.Employee> list9 = new List<Model.Employee>();

                if (!string.IsNullOrEmpty(this.txt_ExcelPath1.Text))
                {
                    foreach (var item in this.txt_ExcelPath1.Text.Split(','))
                    {
                        List<Model.Employee> list = eo.GetExcelDate_Read1(item);
                        list10 = list10.Union(list).ToList();
                    }
                }
                if (!string.IsNullOrEmpty(this.txt_ExcelPath2.Text))
                    list9 = eo.GetExcelDate_Read2(this.txt_ExcelPath2.Text);

                List<Model.Employee> empList = list10.Union(list9).ToList();
                var weekMapping = Common.WeekMapping;

                string outputFileName = Path.Combine(Path.GetDirectoryName(string.IsNullOrEmpty(this.txt_ExcelPath2.Text) ? this.txt_ExcelPath1.Text.Split(',')[0] : this.txt_ExcelPath2.Text), $"{Common.Month}月份考勤总表.xlsx");
                File.Copy(Common.TemplateFileName, outputFileName);
                eo.WriteExcel(outputFileName, weekMapping, empList);

                MessageBox.Show($"导出成功，保存在\"{outputFileName}\"", "提示", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK);
            }

        }
    }
}
