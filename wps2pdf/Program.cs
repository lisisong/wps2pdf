using System;
using System.IO;

namespace WpsToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // 显示Logo
            Version();

            // 如果不带参数，输出帮助信息
            if (args.Length == 0)
            {
                Help();
                Environment.Exit(9);
                return;
            }

            // 判断第1个参数是否-v或-h，如果是，输出相应的信息
            switch (args[0].ToLower().Substring(0, 2))
            {
                case "-v":
                    Environment.Exit(0);
                    return;
                case "-h":
                    Help();
                    Environment.Exit(0);
                    return;
            }

            // 解析文件名
            string wpsFilename = null;
            string pdfFilename = null;
            try
            {
                wpsFilename = Path.GetFullPath(args[0]);
                if (args.Length > 1) { pdfFilename = Path.GetFullPath(args[1]); }
            }
            catch (Exception e)
            {
                Console.WriteLine("参数中包含不正确的文件名");
                Environment.Exit(2);
                return;
            }

            // 判断输入文件是否存在
            if (!File.Exists(wpsFilename))
            {
                Console.WriteLine("错误：指定文件不存在");
                Environment.Exit(1);
                return;
            }

            // 转换
            int exitCode = 0;
            Wps2Pdf wps2pdf = null;
            try
            {
                wps2pdf = new Wps2Pdf();
                wps2pdf.ToPdf(wpsFilename, pdfFilename);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                exitCode = 13;
            }
            finally
            {
                // 不管转换是否成功都退出WPS
                if (wps2pdf != null) { wps2pdf.Dispose(); }
            }

            if (exitCode != 0) Environment.Exit(exitCode);
        }

        static void Version()
        {
            Console.WriteLine(@"wps2pdf - 将WPS文档(含DOC/DOCX)转换为PDF
Copyright (c) 2012 FancyIdea
版本：2.0 (WPS API V9)
");
        }

        static void Help()
        {
            Console.WriteLine(@"
命令：wps2pdf WPS文件 [PDF文件]
      将指定的WPS文件转换为PDF文件，若未指定PDF文件，
      生成的PDF文件与WPS文件同名，且扩展名改为PDF。
");
        }
    }
}
