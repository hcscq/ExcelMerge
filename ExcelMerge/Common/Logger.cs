using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge.Common
{
    public static class Logger
    {
        private static string path = @".\Log\"; //Application.StartupPath + @"\Log\" ;
        private static string fullName = path + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
        public static void WriteLog(string log, bool insertTime = true)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            if (!File.Exists(fullName))
            {
                File.Create(fullName).Close();
            }
            using (FileStream fs = new FileStream(fullName, FileMode.Append))
            {
                using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.GetEncoding("gb2312")))
                {
                    StringBuilder sb = new StringBuilder();
                    if (insertTime)
                        sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append(":");
                    sb.Append(log);
                    sw.Write(sb.ToString());
                    sw.Close();
                }
            }
        }
        public static void WriteHeader(int len = 30, char splitChar = '-', bool timeTitle = false)
        {
            StringBuilder sb = new StringBuilder();
            if (!timeTitle)
                for (int i = 0; i < len; i++)
                {
                    sb.Append(splitChar);
                }
            else
            {
                for (int i = 0; i < len / 2; i++)
                {
                    sb.Append(splitChar);
                }
                sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm")).Append("BEGIN");
                for (int i = 0; i < len / 2; i++)
                {
                    sb.Append(splitChar);
                }
            }
            sb.Append("\r\n");
            WriteLog(sb.ToString());
        }
        public static void WriteTail(int len = 30, char splitChar = '-', bool timeTitle = false)
        {
            StringBuilder sb = new StringBuilder();
            if (!timeTitle)
                for (int i = 0; i < len; i++)
                {
                    sb.Append(splitChar);
                }
            else
            {
                for (int i = 0; i < len / 2; i++)
                {
                    sb.Append(splitChar);
                }
                sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm")).Append("END");
                for (int i = 0; i < len / 2; i++)
                {
                    sb.Append(splitChar);
                }
            }
            sb.Append("\r\n");
            WriteLog(sb.ToString());
        }
    }
}
