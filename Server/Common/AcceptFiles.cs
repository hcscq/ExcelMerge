using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Net.Sockets;
using System.Net;
//using System.Windows.Forms;
using System.IO;
using ExcelMerge.Common;

namespace BusinessLogicLayer
{
    public class ReceiveFiles
    {
        private static Thread threadWatch = null;
        private static Socket socketWatch = null;
        //private static ListBox lstbxMsgView;//显示接受的文件等信息
        //private static ListBox listbOnline;//显示用户连接列表
        private static bool Working = true;
        private static Dictionary<string, Socket> dict = new Dictionary<string, Socket>();
        private static Dictionary<Socket, List<string>> clientFiles = new Dictionary<Socket, List<string>>();
        /// <summary>
        /// 开始监听
        /// </summary>
        /// <param name="localIp"></param>
        /// <param name="localPort"></param>
        public static Exception BeginListening(string localIp, string localPort)
        {
            try
            {
                //创建服务端负责监听的套接字，参数（使用IPV4协议，使用流式连接，使用Tcp协议传输数据）
                socketWatch = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                //获取Ip地址对象
                IPAddress address = IPAddress.Parse(localIp);
                //创建包含Ip和port的网络节点对象
                IPEndPoint endpoint = new IPEndPoint(address, int.Parse(localPort));
                //将负责监听的套接字绑定到唯一的Ip和端口上
                socketWatch.Bind(endpoint);
                //设置监听队列的长度
                socketWatch.Listen(10);
                //创建负责监听的线程，并传入监听方法
                threadWatch = new Thread(WatchConnecting);
                threadWatch.IsBackground = true;//设置为后台线程
                threadWatch.Start();//开始线程
                                    //ShowMgs("服务器启动监听成功");
                                    //ShwMsgForView.ShwMsgforView(lstbxMsgView, "服务器启动监听成功");
                Console.WriteLine("服务器启动监听成功!");
                return null;
            }
            catch (Exception e1)
            {
                return e1;
            }
        }

        /// <summary>
        /// 连接客户端
        /// </summary>
        private static void WatchConnecting()
        {
            while (true&&Working)//持续不断的监听客户端的请求
            {
                //开始监听 客户端连接请求，注意：Accept方法，会阻断当前的线程
                Socket connection = socketWatch.Accept();
                if (connection.Connected&&Working)
                {
                    //向列表控件中添加一个客户端的Ip和端口，作为发送时客户的唯一标识
                    //listbOnline.Items.Add(connection.RemoteEndPoint.ToString());
                    //将与客户端通信的套接字对象connection添加到键值对集合中，并以客户端Ip做为健
                    dict.Add(connection.RemoteEndPoint.ToString(), connection);

                    //创建通信线程
                    ParameterizedThreadStart pts = new ParameterizedThreadStart(RecMsg);
                    Thread thradRecMsg = new Thread(pts);
                    thradRecMsg.IsBackground = true;
                    thradRecMsg.Start(connection);
                    //ShwMsgForView.ShwMsgforView(lstbxMsgView, "客户端连接成功" + connection.RemoteEndPoint.ToString());
                    Console.WriteLine("客户端连接成功" + connection.RemoteEndPoint.ToString());
                }
            }
        }

        /// <summary>
        /// 接收消息
        /// </summary>
        /// <param name="socketClientPara"></param>
        private static void RecMsg(object socketClientPara)
        {
            Socket socketClient = socketClientPara as Socket;

            while (true)
            {
                //定义一个接受用的缓存区（100M字节数组）
                //byte[] arrMsgRec = new byte[1024 * 1024 * 100];
                //将接收到的数据存入arrMsgRec数组,并返回真正接受到的数据的长度   
                if (socketClient.Connected)
                {
                    try
                    {
                        //因为终端每次发送文件的最大缓冲区是512字节，所以每次接收也是定义为512字节
                        byte[] buffer = new byte[Setting.MaxBuffLength];
                        int size = 0;
                        long len = 0;
                        string fileSavePath = @".\\files\\";//获得用户保存文件的路径

                        string fileName=string.Empty;
                        string fileEx = string.Empty;
                        //创建文件流，然后让文件流来根据路径创建一个文件
                        FileStream fs = null;
                        //从终端不停的接受数据，然后写入文件里面，只到接受到的数据为0为止，则中断连接

                        DateTime oTimeBegin = DateTime.Now;

                        int offset = 0;
                        int Mode = 0;//0:wait file header 1:wait data
                        long fileLen = 0;
                        ClientPacketId Key = 0;
                        clientFiles.Add(socketClient,new List<string>());
                        
                        while ((size = socketClient.Receive(buffer, 0, buffer.Length, SocketFlags.None)) > 0)
                        {
                            if (Mode == 0) Key = ClientPacketId.DataWithFileName;
                            else if (Mode == 1)
                            {
                                if (len + size >= fileLen)
                                    Key = ClientPacketId.FileEnd;
                                else
                                    Key = ClientPacketId.FileData;
                            }
                            switch (Key)
                            {
                                case ClientPacketId.FileData:
                                    offset = 0;
                                    break;
                                case ClientPacketId.FileName:
                                case ClientPacketId.DataWithFileName: 
                                    using (Stream stream = new MemoryStream(buffer))
                                    using (BinaryReader br = new BinaryReader(stream))
                                    {
                                        br.ReadByte();
                                        fileName = br.ReadString();
                                        fileEx = br.ReadString();
                                        fileLen = br.ReadInt64();
                                        offset = (int)stream.Position;
                                    }

                                    if (!string.IsNullOrEmpty(fileName))
                                    {
                                        if (!Directory.Exists(fileSavePath))
                                            Directory.CreateDirectory(fileSavePath);
                                        fileName = fileSavePath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + fileName+"."+fileEx;
                                        fs = new FileStream(fileName, FileMode.Create);
                                        Mode = 1;
                                    }
                                    break;
                                case ClientPacketId.FileEnd:
                                    fs.Write(buffer, offset,(int)(fileLen-len));
                                    fs.Flush();
                                    fs.Close();
                                    clientFiles[socketClient].Add(fileName);
                                    Console.WriteLine("文件保存成功:" + fileName);
                                    fileName = string.Empty;
                                    fileEx = string.Empty;
                                    Mode = 0;
                                    len = 0;
                                    fileLen = 0;
                                    continue;
                                default:
                                    break;
                            }

                            fs.Write(buffer, offset, size-offset);
                            len += size;
                        }
                        DateTime oTimeEnd = DateTime.Now;
                        TimeSpan oTime = oTimeEnd.Subtract(oTimeBegin);


                        Console.WriteLine(socketClient.RemoteEndPoint + "断开连接");
                        dict.Remove(socketClient.RemoteEndPoint.ToString());

                        socketClient.Close();

                        //ShwMsgForView.ShwMsgforView(lstbxMsgView, "文件保存成功:" + fileName);
                        //ShwMsgForView.ShwMsgforView(lstbxMsgView, "接收文件用时:" + oTime.ToString() + ",文件大小：" + len / 1024 + "kb");
                    }
                    catch(Exception e1)
                    {
                        Console.WriteLine(socketClient.RemoteEndPoint + "下线了");
                        clientFiles.Remove(socketClient);
                        dict.Remove(socketClient.RemoteEndPoint.ToString());

                        break;
                    }
                }
                else
                {

                }
            }
        }

        /// <summary>
        /// 关闭连接
        /// </summary>
        public static void CloseTcpSocket()
        {
            dict.Clear();
            //listbOnline.Items.Clear();
            threadWatch.Abort();
            socketWatch.Close();
            //ShwMsgForView.ShwMsgforView(lstbxMsgView, "服务器关闭监听");
        }
    }


}