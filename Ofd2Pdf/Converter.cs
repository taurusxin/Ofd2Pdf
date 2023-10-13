using Spire.Pdf.Conversion;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ofd2Pdf
{
    public enum ConvertResult
    {
        Successful,
        Failed
    }
    public class Converter
    {
        public ConvertResult ConvertToPdf(string Input, string OutPut)
        {
            Console.WriteLine(Input + " " + OutPut);
            if (Input == null || OutPut == null)
            {
                return ConvertResult.Failed;
            }

            if (!File.Exists(Input))
            {
                return ConvertResult.Failed;
            }

            try
            {
                OfdConverter converter = new OfdConverter(Input);
                converter.ToPdf(OutPut);
                return ConvertResult.Successful;
            }
            catch (Exception)
            {
                return ConvertResult.Failed;
            }
        }
    }
}
