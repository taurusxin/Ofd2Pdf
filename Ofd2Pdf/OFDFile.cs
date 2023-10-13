using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ofd2Pdf
{
    public enum Status
    {
        等待转换,
        正在转换,
        转换完成,
        转换失败
    }
    internal class OFDFile
    {
        public string FileName { get; set; }
        public Status Status { get; set; }

        public OFDFile(string fileName, Status status)
        {
            FileName = fileName;
            Status = status;
        }

        public OFDFile(string fileName)
        {
            FileName = fileName;
            Status = Status.等待转换;
        }
    }
}
