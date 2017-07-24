using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyCharRoomClient;
using System.Net;
using System.Net.Sockets;
using ExcelMerge.Common;
namespace Client
{

    public partial class Form1 : Form
    {
        private string Str_MainExcelPath = Application.StartupPath + "\\Main.xlsx";
        private string Str_CurChildExcelPath = Application.StartupPath;
        private string Str_ChildrenDir = Application.StartupPath + "\\Children";

        private List<string> ChildrenFilesName = new List<string>();

        public enum OperRe { ALLRIGHT, SKIPED, PART, ERROR, RETRY }

        Socket SerSocket = null;
        public Form1()
        {
            InitializeComponent();
            Inital();
            DoWork();
        }
        public void DoWork()
        {
            IPEndPoint remotePoint = new IPEndPoint(IPAddress.Parse(Setting.IP), int.Parse(Setting.Port));
            SocketAsyncEventArgs socketArg = new SocketAsyncEventArgs();
            socketArg.RemoteEndPoint = remotePoint;
            socketArg.Completed += SocketArg_Completed;
            Socket.ConnectAsync(SocketType.Stream, ProtocolType.Tcp,socketArg);
        }

        private void SocketArg_Completed(object sender, SocketAsyncEventArgs e)
        {
            if (e.SocketError!=SocketError.Success)
            {
                MessageBox.Show("Server offline.Please wait a momment and try again.");
                return;
            }
            SerSocket = e.ConnectSocket;
            if (SerSocket.Connected)
            {
                for (int i = 0; i < ChildrenFilesName.Count; i++)
                {
                    Net.SendFile(SerSocket,ChildrenFilesName[i],Setting.MaxBuffLength,Setting.OutTime);
                }
                MessageBox.Show("All sended,please wait server return.");
            }
        }

        public void Inital()
        {
            this.Text = "Excel合并";
            //if (!File.Exists(Str_MainExcelPath))
            //{
            //    MessageBox.Show("在此程序当前目录下放主Excel,名字为：Main.xlsx（注意是XLSX）！");
            //    return;
            //}

            if (!Directory.Exists(Str_ChildrenDir))
            {
                Directory.CreateDirectory(Str_ChildrenDir);
                MessageBox.Show("在此程序当前目录下的Children文件夹中放子Excel(注意是XLS 2003格式的).");
                return;
            }
            ChildrenFilesName = Directory.GetFiles(Str_ChildrenDir, "*.xls").ToList();

            ChildrenFilesName.AddRange(Directory.GetFiles(Str_ChildrenDir, "*.xlsx").ToList());
        }
    }

}
