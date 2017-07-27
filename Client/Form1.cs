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
using SendCommon;
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
        private FileStream fs = null;
        private byte Mode = 0;//0:wait fileName 1 Wait Data
        private long fileLen = 0;
        private long ReadLen = 0;

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
            SocketAsyncEventArgs connectedArg = new SocketAsyncEventArgs();
            connectedArg.RemoteEndPoint = remotePoint;
            connectedArg.Completed += ConnectedArg_Completed;
            Socket.ConnectAsync(SocketType.Stream, ProtocolType.Tcp,connectedArg);
        }

        private void ConnectedArg_Completed(object sender, SocketAsyncEventArgs e)
        {
            if (e.SocketError!=SocketError.Success)
            {
                MessageBox.Show("Server offline.Please wait a momment and try again.");
                return;
            }
            SerSocket = e.ConnectSocket;
            if (SerSocket.Connected)
            {
                SocketAsyncEventArgs asyncArg = new SocketAsyncEventArgs();
                if (asyncArg == null) throw new Exception("Not enough BUFF.");
                asyncArg.SetBuffer(new byte[Setting.MaxBuffLength], 0, Setting.MaxBuffLength);
                asyncArg.Completed += AsyncArg_Completed;
                using (Stream stream = new MemoryStream(asyncArg.Buffer))
                using (BinaryWriter br = new BinaryWriter(stream))
                {
                    br.Write((byte)ClientPacketId.FileCount);
                    br.Write(ChildrenFilesName.Count);
                }
                if (!SerSocket.SendAsync(asyncArg))
                    AsyncArg_Completed(SerSocket,asyncArg);

            }
        }

        private void AsyncArg_Completed(object sender, SocketAsyncEventArgs e)
        {
            for (int i = 0; i < ChildrenFilesName.Count; i++)
                Net.SendFile(SerSocket, ChildrenFilesName[i], Setting.MaxBuffLength, Setting.OutTime);

            MessageBox.Show("All sended,please wait server return.");
            SocketAsyncEventArgs asyncArg = BufferManager.GetBuffer();
            asyncArg.Completed += new EventHandler<SocketAsyncEventArgs>(IOCompleted);
            //asyncArg.SetBuffer(new byte[Setting.MaxBuffLength],0);
            if (!SerSocket.ReceiveAsync(asyncArg))
                IOCompleted(SerSocket,asyncArg);
        }

        private void IOCompleted(object sender,SocketAsyncEventArgs arg)
        {
            if (arg.SocketError == SocketError.Success&&arg.BytesTransferred>0)
            {
                string fileName = string.Empty;
                string fileExtension = string.Empty;
                string fileSavePath = ".\\";

                int offSet = 0;
                if (Mode == 0)
                {
                    using (Stream stream = new MemoryStream(arg.Buffer))
                    using (BinaryReader sr = new BinaryReader(stream))
                    {
                        sr.ReadByte();
                        fileName = sr.ReadString();
                        fileExtension = sr.ReadString();
                        fileLen = sr.ReadInt64();
                        if (!string.IsNullOrEmpty(fileName))
                        {
                            if (!Directory.Exists(fileSavePath))
                                Directory.CreateDirectory(fileSavePath);
                            fileName = fileSavePath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + fileName + "." + fileExtension;
                            fs = new FileStream(fileName, FileMode.Create);
                            Mode = 1;
                        }
                        offSet = (int)stream.Position;
                    }
                }
                ReadLen += arg.BytesTransferred;
                if (ReadLen < fileLen)
                {
                    fs.Write(arg.Buffer, offSet, arg.BytesTransferred - offSet);

                    if (!((Socket)sender).ReceiveAsync(arg))
                        IOCompleted(sender,arg);
                }
                else
                {
                    fs.Write(arg.Buffer, offSet,(int)(fileLen+arg.BytesTransferred-ReadLen - offSet));
                    BufferManager.ReleaseBuffer(arg);
                    //fs.Flush();
                    fs.Close();
                    fs = null;
                    Mode = 0;
                    ReadLen = 0;
                    fileLen = 0;
                }
            }
            else
            {
                if (fs != null) { fs.Close();fs = null; }
                ((Socket)sender).Close();
                BufferManager.ReleaseBuffer(arg);
                Mode = 0;
                ReadLen = 0;
                fileLen = 0;
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
            ChildrenFilesName = Directory.GetFiles(Str_ChildrenDir, "*.xls?").ToList();

            //ChildrenFilesName.AddRange(Directory.GetFiles(Str_ChildrenDir, "*.xlsx").ToList());
        }
    }

}
