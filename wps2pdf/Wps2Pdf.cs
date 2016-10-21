using System;
using System.IO;
using Word;

namespace WpsToPdf
{
    class Wps2Pdf : IDisposable
    {
        dynamic wps;

        public Wps2Pdf()
        {
            Type type = Type.GetTypeFromProgID("KWps.Application");
            wps = Activator.CreateInstance(type);
        }
        public void ToPdf(string wpsFilename, string pdfFilename = null)
        {
            if (wpsFilename == null) { throw new ArgumentNullException("wpsFilename"); }

            if (pdfFilename == null)
            {
                pdfFilename = Path.ChangeExtension(wpsFilename, "pdf");
            }

            Console.WriteLine(string.Format(@"正在转换 [{0}]
      -> [{1}]", wpsFilename, pdfFilename));

            dynamic doc = wps.Documents.Open(wpsFilename, Visible: false);
            doc.ExportAsFixedFormat(pdfFilename, WdExportFormat.wdExportFormatPDF);
            doc.Close();
        }

        public void Dispose()
        {
            if (wps != null) { wps.Quit(); }
        }
    }
}
