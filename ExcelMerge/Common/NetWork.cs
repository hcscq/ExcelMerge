﻿using System;
using System.Net;
using System.Net.Sockets;
using System.IO;
using ExcelMerge.Common;
namespace SendCommon
{

    /// <summary>
    /// Net : 提供静态方法，对常用的网络操作进行封装
    /// </summary>
    public sealed class Net
    {
        private Net()
        {
        }

        /// <summary>
        /// 向远程主机发送数据
        /// </summary>
        /// <param name="socket">要发送数据且已经连接到远程主机的 Socket</param>
        /// <param name="buffer">待发送的数据</param>
        /// <param name="outTime">发送数据的超时时间，以秒为单位，可以精确到微秒</param>
        /// <returns>0:发送数据成功；-1:超时；-2:发送数据出现错误；-3:发送数据时出现异常</returns>
        /// <remarks >
        /// 当 outTime 指定为-1时，将一直等待直到有数据需要发送
        /// </remarks>
        public static int SendData(Socket socket, byte[] buffer, int outTime)
        {
            if (buffer.Length != Setting.MaxBuffLength)
                throw new Exception("Length error");
            if (socket.Send(buffer) == buffer.Length)
                return 0;
            else throw new Exception("Send too less.");
            if (socket == null || socket.Connected == false)
            {
                throw new ArgumentException("参数socket 为null，或者未连接到远程计算机");
            }
            if (buffer == null || buffer.Length == 0)
            {
                throw new ArgumentException("参数buffer 为null ,或者长度为 0");
            }

            int flag = 0;
            try
            {
                int left = buffer.Length;
                int sndLen = 0;

                while (true)
                {
                    if ((socket.Poll(outTime * 100, SelectMode.SelectWrite) == true))
                    {        // 收集了足够多的传出数据后开始发送
                        sndLen = socket.Send(buffer, sndLen, left, SocketFlags.None);
                        left -= sndLen;
                        if (left == 0)
                        {                                        // 数据已经全部发送
                            flag = 0;
                            break;
                        }
                        else
                        {
                            if (sndLen > 0)
                            {                                    // 数据部分已经被发送
                                continue;
                            }
                            else
                            {                                                // 发送数据发生错误
                                flag = -2;
                                break;
                            }
                        }
                    }
                    else
                    {                                                        // 超时退出
                        flag = -1;
                        break;
                    }
                }
            }
            catch (SocketException e)
            {

                flag = -3;
            }
            return flag;
        }


        /// <summary>
        /// 向远程主机发送文件
        /// </summary>
        /// <param name="socket" >要发送数据且已经连接到远程主机的 socket</param>
        /// <param name="fileName">待发送的文件名称</param>
        /// <param name="maxBufferLength">文件发送时的缓冲区大小</param>
        /// <param name="outTime">发送缓冲区中的数据的超时时间</param>
        /// <returns>0:发送文件成功；-1:超时；-2:发送文件出现错误；-3:发送文件出现异常；-4:读取待发送文件发生错误</returns>
        /// <remarks >
        /// 当 outTime 指定为-1时，将一直等待直到有数据需要发送
        /// </remarks>
        public static int SendFile(Socket socket, string fileName, int maxBufferLength, int outTime)
        {
            if (fileName == null || maxBufferLength <= 0)
            {
                throw new ArgumentException("待发送的文件名称为空或发送缓冲区的大小设置不正确.");
            }
            int flag = 0;
            try
            {
                FileInfo fileInfo = new FileInfo(fileName);

                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                long fileLen = fs.Length;                        // 文件长度
                long leftLen = fileLen;                            // 未读取部分
                int readLen = 0;                                // 已读取部分
                int offSet = 0;
                byte[] buffer = null;
                byte[] bufferSend = new byte[maxBufferLength];
                using (Stream stream = new MemoryStream(bufferSend))
                using (BinaryWriter bw = new BinaryWriter(stream))
                {
                    bw.Write((byte)ClientPacketId.DataWithFileName);
                    bw.Write(Path.GetFileNameWithoutExtension(fileInfo.Name));
                    bw.Write(fileInfo.Extension);
                    bw.Write(stream.Position + fileLen + 8);
                    offSet = (int)stream.Position;
                }


                if (fileLen+offSet <= maxBufferLength)
                {            /* 文件可以一次读取*/
                    buffer = new byte[maxBufferLength];

                    Array.Copy(bufferSend, buffer, offSet);

                    readLen = fs.Read(bufferSend, 0, (int)fileLen);  
                    Array.Copy(bufferSend,0, buffer,offSet, readLen);
                    flag = SendData(socket, buffer, outTime);
                }
                else
                {
                    /* 循环读取文件,并发送 */
                    buffer = new byte[maxBufferLength - offSet];

                    readLen = fs.Read(buffer, 0, buffer.Length);

                    Array.Copy(buffer, 0, bufferSend, offSet, readLen);

                    if ((flag = SendData(socket, bufferSend, outTime)) < 0)
                    {
                        throw new Exception("Send 0 byte.");
                    }

                    leftLen -= readLen;
                    while (leftLen > 0)
                    {
                        buffer = new byte[maxBufferLength];
                        if (leftLen < maxBufferLength)
                        {
                            //buffer = new byte[maxBufferLength];
                            //buffer[0] = (byte)ClientPacketId.FileData;
                            readLen = fs.Read(buffer, 0, Convert.ToInt32(leftLen));
                        }
                        else
                        {
                            //buffer = new byte[maxBufferLength];
                            //buffer[0] = (byte)ClientPacketId.FileData;
                            readLen = fs.Read(buffer, 0, maxBufferLength);
                        }
                        if ((flag = SendData(socket, buffer, outTime)) < 0)
                        {
                            break;
                        }
                        leftLen -= readLen;
                    }
                }

                //fs.Flush();
                fs.Close();
            }
            catch (IOException e)
            {

                flag = -4;
                throw new Exception("Not deal.");
            }
            return flag;
        }

    }
}