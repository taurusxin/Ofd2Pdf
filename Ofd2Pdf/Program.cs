using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ofd2Pdf
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            else
            {
                Converter converter = new Converter();

                bool hasFailed = false;

                for (int i = 0; i < args.Length; i++)
                {
                    string file = args[i];
                    string PdfName = file.Substring(0, file.Length - 3) + "pdf";

                    var result = converter.ConvertToPdf(file, PdfName);

                    if (result == ConvertResult.Failed)
                    {
                        Console.WriteLine("[Failed]: " + file);
                        hasFailed = true;
                    } else
                    {
                        Console.WriteLine("[Success]: " + file);
                    }
                }

                if (hasFailed)
                {
                    Environment.Exit(1);
                }
                else
                {
                    Environment.Exit(0);
                }
            }
        }
    }
}
