using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Utils;
using System.Data.OleDb;
using System.IO;
namespace WindowsFormsApplication1
{
    public partial class ExcelMerge : Form
    {
        private string Str_MainExcelPath = Application.StartupPath + "\\Main.xlsx";
        private string Str_CurChildExcelPath = Application.StartupPath;
        private string Str_ChildrenDir = Application.StartupPath + "\\Children";
        private OleDbConnection Con_MainExcel = null;
        private OleDbConnection Con_CurChildExcel = null;
        private List<string> ChildSheetName = new List<string>();
        private List<string> ChildrenFilesName = new List<string>();
        public ExcelMerge()
        {
            InitializeComponent();
            Inital();
            MergeExcel();
            MessageBox.Show("处理完成.");
        }
        public void Inital()
        {
            this.Text = "Excel合并";
            if (!File.Exists(Str_MainExcelPath)) {
                MessageBox.Show("在此程序当前目录下放主Excel,名字为：Main.xlsx（注意是XLSX）！");
                return;
            }
                Con_MainExcel = ExcelHelper.CreateConnection(Str_MainExcelPath, ExcelHelper.ExcelVerion.Excel2007);
            if (!Directory.Exists(Str_ChildrenDir))
            {
                Directory.CreateDirectory(Str_ChildrenDir);
                MessageBox.Show("在此程序当前目录下的Children文件夹中放子Excel(注意是XLS 2003格式的).");
                return;
            }
            ChildrenFilesName=Directory.GetFiles(Str_ChildrenDir, "*.xls").ToList();
        }
        public void MergeExcel()
        {
            DataTable mainDt = ExcelHelper.ExecuteDataTable(Con_MainExcel, "select * from [Sheet1$]", null);
            bool retry = false;
            string fn = "";
            for (int m = 0; m < ChildrenFilesName.Count; m++)
            {
                ChildSheetName.Clear();
                try
                {
                    fn = ChildrenFilesName[m];
                    if (!retry)
                        Con_CurChildExcel = ExcelHelper.CreateConnection(fn, ExcelHelper.ExcelVerion.Excel2003);
                    Con_CurChildExcel.Open();
                    foreach (DataRow drt in Con_CurChildExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }).Rows)
                        ChildSheetName.Add(drt["TABLE_NAME"].ToString());
                    Con_CurChildExcel.Close();
                    DataTable dt;
                    DataRow dr;
                    foreach (var it in ChildSheetName)
                    {
                        if (it.Contains("_"))
                        {
                            MessageBox.Show("检查文件" + fn + "\r\n是否启用了筛选功能.目前将跳过该文件的工作簿:" + it + ".稍后请自行检查该部分数据内容的准确性.");
                            continue;
                        }
                        dt = ExcelHelper.ExecuteDataTable(Con_CurChildExcel, @"select * from [" + it + "] ", null);
                        if (dt.Rows.Count > 2)
                            for (int i = 2; i < dt.Rows.Count; i++)
                            {
                                if (dt.Rows[i]["F4"] == null || string.IsNullOrWhiteSpace(dt.Rows[i]["F4"].ToString()))
                                    continue;
                                dr = mainDt.NewRow();
                                dr["商务代表"] = dt.Rows[0]["F2"];
                                dr["上级经理"] = dt.Rows[0]["F4"];
                                dr["医院名称"] = dt.Rows[i]["F4"] == null ? " " : dt.Rows[i]["F4"];
                                dr["医生姓名"] = dt.Rows[i]["F3"] == null ? " " : dt.Rows[i]["F3"];
                                dr["科室"] = dt.Rows[i]["F5"] == null ? " " : dt.Rows[i]["F5"];
                                dr["客户性质"] = dt.Rows[i]["F6"] == null ? " " : dt.Rows[i]["F6"];
                                dr["当日情况反馈"] = dt.Rows[i]["F7"] == null ? " " : dt.Rows[i]["F7"];
                                dr["计划"] = dt.Rows[i]["F8"] == null ? " " : dt.Rows[i]["F8"];
                                dr["省区意见"] = dt.Rows[i]["F9"] == null ? " " : dt.Rows[i]["F9"];
                                mainDt.Rows.Add(dr);
                                //sssss = GetInsertStr(dr);
                                ExcelHelper.ExecuteNonQuery(Con_MainExcel, GetInsertStr(dr), null);
                            }
                    }
                    retry = false;
                }
                catch (Exception e1)
                {
                    //预期的格式
                    if (e1.Message.Contains("预期的格式") && !retry)
                    {
                        retry = true;
                        m--;
                        Con_CurChildExcel = ExcelHelper.CreateConnection(fn, ExcelHelper.ExcelVerion.Excel2007);
                    }
                    else
                    {
                        retry = false;
                        MessageBox.Show("文件" + fn + "处理出现问题.截图发给..李大爷\r\n" + e1.Message);
                    }
                }
            }
            return;
        }
        private string GetInsertStr(DataRow dr, string tableName = " [Sheet1$] ")
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("insert into ").Append(tableName).Append("( ");
            foreach (var it in dr.Table.Columns)
            {
                if (dr[it.ToString()] != null&& !string.IsNullOrWhiteSpace(dr[it.ToString()].ToString()))
                    sb.Append("[").Append(it.ToString()).Append("],");
            }
            sb.Remove(sb.Length - 1, 1).Append(")").Append("values(");
            foreach (var it in dr.Table.Columns)
            {
                if (dr[it.ToString()] != null && !string.IsNullOrWhiteSpace(dr[it.ToString()].ToString()))
                    sb.Append("'").Append(dr[it.ToString()].ToString().Replace('\r',' ').Replace('\n',' ')).Append("',");
            }
            sb.Remove(sb.Length - 1, 1).Append(")");
            return sb.ToString();
        }
    }
}
