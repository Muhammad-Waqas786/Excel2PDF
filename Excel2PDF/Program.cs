using Excel2PDF.Core;
using System.Configuration;
using System.IO;
using Excel2PDF.ExcelProcessor;
using Excel2PDF.PDFProcessor;

namespace Excel2PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            IExcelProcessor excelProcessor = new ExcelProcessor.ExcelProcessor();
            IPDFProcessor pdfProcessor = new PDFProcessor.PDFProcessor();
            var config = GetConfig();

            var excelReadFolder = new DirectoryInfo(config.ExcelReadFolder);

            foreach (var file in excelReadFolder.GetFiles())
            {
                string fileExtension = file.Name.Substring(file.Name.LastIndexOf("."));
                string newFileName = file.Name.Replace(fileExtension, "_intro.pdf");

                using (var stream = File.Open(file.FullName, FileMode.Open, FileAccess.Read))
                {
                    var infoTemp = excelProcessor.ParseExcelForInfo(stream);
                    pdfProcessor.GenerateIntroPDF(config, infoTemp, newFileName);
                }

                File.Move(file.FullName, $"{config.ExcelArchiveFolder}\\{file.Name}");
            }
        }

        static Excel2PDFConfig GetConfig()
        {
            return new Excel2PDFConfig
            {
                PDFTemplateFolder = ConfigurationManager.AppSettings["PDFTemplateFolder"],
                PDFWriteFolder = ConfigurationManager.AppSettings["PDFWriteFolder"],
                ExcelReadFolder = ConfigurationManager.AppSettings["ExcelReadFolder"],
                DateFormat = ConfigurationManager.AppSettings["DateFormat"],
                ExcelArchiveFolder = ConfigurationManager.AppSettings["ExcelArchiveFolder"]
            };
        }
    }
}
