using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge.Common
{
    enum ClientPacketId : byte
    {
        FileName,
        FileData,
        DataWithFileName,
        FileEnd,
        FileCount,
        MAX=FileEnd
    }
    enum ServerPacketId : byte
    {
        IsLastFile,
        ReFileName,
        ReFileExtension
    }
    public static class Setting
    {
        public const string IP = "127.0.0.1";
        public const string Port = "8012";
        public const int MaxBuffLength = 512;
        public const int OutTime = 5000;
        public const int MaxBufferCount = 20;
    }
    public class SocketBuffer
    {
        public SocketAsyncEventArgs SocketArg = new SocketAsyncEventArgs();
        public bool IsEmpty = true;
        public  SocketBuffer()
        {
            SocketArg.SetBuffer(new byte[Setting.MaxBuffLength], 0, Setting.MaxBuffLength);
            SocketArg.UserToken = this;
        }
    }
    public  class BufferManager
    {
        private static List<SocketBuffer> BufferList = new List<SocketBuffer>();
        public static SocketAsyncEventArgs GetBuffer()
        {
            for (int i = 0; i < BufferList.Count; i++)
            {
                if (BufferList[i].IsEmpty)
                {
                    BufferList[i].IsEmpty = false;
                    return BufferList[i].SocketArg;
                }
            }
            if (BufferList.Count < Setting.MaxBufferCount)
            {
                SocketBuffer sb = new SocketBuffer();
                BufferList.Add(sb);
                sb.IsEmpty = false;
                return sb.SocketArg;
            }
            else return null;
        }
        public static void ReleaseBuffer(SocketAsyncEventArgs socketArg)
        {
            ((SocketBuffer)socketArg.UserToken).IsEmpty = true; ;
        }
    }
    class Common
    {
    }
}
