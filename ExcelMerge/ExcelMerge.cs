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
using ExcelMerge.Common;

namespace ExcelMerge
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
        private List<sheetInfo> ListSheetInfo = new List<sheetInfo>();
        public enum OperRe{ALLRIGHT,SKIPED,PART,ERROR,RETRY }

        public struct sheetInfo
        {
            public string fileName;
            public string sheetName;
            public OperRe operRe;
            public int dealCount;
            public override string ToString()
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(" FileName:").Append(fileName);
                sb.Append(" SheetName:").Append(sheetName);
                sb.Append(" DealCount:").Append(dealCount);
                sb.Append(" OperRe:");
                switch (operRe)
                {
                    case OperRe.ALLRIGHT:
                        sb.Append("全部顺利执行.");break;
                    case OperRe.PART:
                        sb.Append("部分执行.");break;
                    case OperRe.SKIPED:
                        sb.Append("被跳过.");break;
                    case OperRe.ERROR:
                        sb.Append("出错并跳过."); break;
                    case OperRe.RETRY:
                        sb.Append("出错，立即重试."); break;
                    default:break;
                }
                return sb.ToString();
            }
        }
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
            Logger.WriteHeader();
            ListSheetInfo.Clear();
            DataTable mainDt = ExcelHelper.ExecuteDataTable(Con_MainExcel, "select * from [Sheet1$]", null);
            bool retry = false;
            string fn = "";
            sheetInfo loc_info=new sheetInfo ();
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
                    string[,] colName = { {"F1","F2","F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14"}
                    , { "所属大区","所属省区","商务代表","上级经理","拜访日期","医院名称","医生姓名","科室","上月订单数","本周当前_订单数",
                            "关注患者数","拜访目的_（需要解决的问题）","当日情况反馈_（需为有价值信息，作为备案便于下次跟进）","改善计划_（准备如何改进）"} };
                    foreach (var it in ChildSheetName)
                    {
                        loc_info = new sheetInfo();
                        loc_info.fileName = fn;
                        loc_info.sheetName = it;
                        if (ListSheetInfo.Exists(item => item.fileName == loc_info.fileName && item.sheetName == loc_info.sheetName))
                            continue;
                        if (it.Contains("_"))
                        {
                            MessageBox.Show("检查文件" + fn + "\r\n是否启用了筛选功能.目前将跳过该文件的工作簿:" + it + ".稍后请自行检查该部分数据内容的准确性.");
                            loc_info.operRe = OperRe.SKIPED;
                            ListSheetInfo.Add(loc_info);
                            continue;
                        }
                        dt = ExcelHelper.ExecuteDataTable(Con_CurChildExcel, @"select * from [" + it + "A:P] ", null);


                        try
                        {
                            int colNameIndex =dt.Columns[0].ColumnName.Equals("F1")? 0:1;
                            
                            if (dt.Rows.Count > 2)
                                for (int i = 1-colNameIndex; i < dt.Rows.Count; i++)
                                {
                                    //if ((dt.Rows[i]["F4"] == null || string.IsNullOrWhiteSpace(dt.Rows[i]["F4"].ToString())) && (dt.Rows[i]["F3"] == null || string.IsNullOrWhiteSpace(dt.Rows[i]["F3"].ToString()))
                                    //    && (dt.Rows[i]["F5"] == null || string.IsNullOrWhiteSpace(dt.Rows[i]["F5"].ToString())) &&
                                    //    (dt.Rows[i]["F6"] == null || string.IsNullOrWhiteSpace(dt.Rows[i]["F6"].ToString())) && (dt.Rows[i]["F7"] == null || string.IsNullOrWhiteSpace(dt.Rows[i]["F7"].ToString())))
                                    //    continue;
                                    dr = mainDt.NewRow();
                                    if (dt.Rows[i][colName[colNameIndex, 0]] == null || string.IsNullOrWhiteSpace(dt.Rows[i][colName[colNameIndex,0]].ToString()))
                                        dr["所属大区"] = dt.Rows[1][colName[colNameIndex, 0]];
                                    else
                                        dr["所属大区"] = dt.Rows[i][colName[colNameIndex, 0]];

                                    if (dt.Rows[i][colName[colNameIndex,1]] == null || string.IsNullOrWhiteSpace(dt.Rows[i][colName[colNameIndex,1]].ToString()))
                                        dr["所属省区"] = dt.Rows[1][colName[colNameIndex,1]];
                                    else dr["所属省区"] = dt.Rows[i][colName[colNameIndex,1]];

                                    if (dt.Rows[i][colName[colNameIndex,2]] == null || string.IsNullOrWhiteSpace(dt.Rows[i][colName[colNameIndex,2]].ToString()))
                                        dr["商务代表"] = dt.Rows[1][colName[colNameIndex,2]];
                                    else dr["商务代表"] = dt.Rows[i][colName[colNameIndex,2]];

                                    if (dt.Rows[i][colName[colNameIndex,3]] == null || string.IsNullOrWhiteSpace(dt.Rows[i][colName[colNameIndex,3]].ToString()))
                                        dr["上级经理"] = dt.Rows[1][colName[colNameIndex,3]];
                                    else dr["上级经理"] = dt.Rows[i][colName[colNameIndex,3]];

                                    dr["拜访日期"] = dt.Rows[i][colName[colNameIndex,4]];
                                    dr["医院名称"] = dt.Rows[i][colName[colNameIndex, 5]];
                                    dr["医生姓名"] = dt.Rows[i][colName[colNameIndex, 6]];
                                    dr["科室"] = dt.Rows[i][colName[colNameIndex, 7]];
                                    dr["上月订单数"] = dt.Rows[i][colName[colNameIndex, 8]];
                                    dr["本周当前_订单数"] = dt.Rows[i][colName[colNameIndex, 9]];
                                    dr["关注患者数"] = dt.Rows[i][colName[colNameIndex, 10]];
                                    dr["拜访目的_（需要解决的问题）"] = dt.Rows[i][colName[colNameIndex, 11]];
                                    dr["当日情况反馈_（需为有价值信息，作为备案便于下次跟进）"] = dt.Rows[i][colName[colNameIndex, 12]];
                                    dr["改善计划_（准备如何改进）"] = dt.Rows[i][colName[colNameIndex, 13]];

                                    //dr["医院名称"] = dt.Rows[i]["F4"] == null ? "未填写" : dt.Rows[i]["F4"];
                                    //dr["医生姓名"] = dt.Rows[i]["F3"] == null ? "未填写" : dt.Rows[i]["F3"];
                                    //dr["科室"] = dt.Rows[i]["F5"] == null ? "未填写" : dt.Rows[i]["F5"];
                                    //dr["客户性质"] = dt.Rows[i]["F6"] == null ? "未填写" : dt.Rows[i]["F6"];
                                    //dr["当日情况反馈"] = dt.Rows[i]["F7"] == null ? "未填写" : dt.Rows[i]["F7"];
                                    //dr["计划"] = dt.Rows[i]["F8"] == null ? " " : dt.Rows[i]["F8"];
                                    //dr["省区意见"] = dt.Rows[i]["F9"] == null ? " " : dt.Rows[i]["F9"];
                                    mainDt.Rows.Add(dr);
                                    //sssss = GetInsertStr(dr);
                                    if (0 < ExcelHelper.ExecuteNonQuery(Con_MainExcel, GetInsertStr(dr), null))
                                        loc_info.dealCount++;
                                }
                        }
                        catch (Exception e1)
                        {
                            loc_info.operRe = OperRe.ERROR;
                            Logger.WriteLog(loc_info.ToString() + e1.Message + "\r\n");
                            retry = false;
                            MessageBox.Show("文件" + fn + "处理出现问题.\r\n工作表名："+it + e1.Message);
                            continue;
                        }
                        loc_info.operRe = OperRe.ALLRIGHT;
                        ListSheetInfo.Add(loc_info);
                    }
                    retry = false;
                }
                catch (Exception e1)
                {
                    //Logger.WriteLog(e1.Message + "\r\n");
                    //预期的格式
                    if (e1.Message.Contains("预期的格式") && !retry)
                    {
                        retry = true;
                        m--;
                        Con_CurChildExcel = ExcelHelper.CreateConnection(fn, ExcelHelper.ExcelVerion.Excel2007);
                    }
                    else
                    {
                        loc_info.operRe = OperRe.ERROR;
                        Logger.WriteLog(loc_info.ToString() + e1.Message + "\r\n");
                        retry = false;
                        MessageBox.Show("文件" + fn + "处理出现问题.截图发给..\r\n" + e1.Message);
                    }
                }
            }
            foreach (var it in ListSheetInfo)
                Logger.WriteLog(it.ToString()+"\r\n");
            Logger.WriteTail();
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
