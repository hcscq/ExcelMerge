using System;
using System.Collections.Generic;
using System.Linq;
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
        MAX=FileEnd
    }
    public static class Setting
    {
        public const string IP = "127.0.0.1";
        public const string Port = "8012";
        public const int MaxBuffLength = 512;
        public const int OutTime = 5000;
    }
    class Common
    {
    }
}
