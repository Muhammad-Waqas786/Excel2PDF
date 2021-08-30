using Excel2PDF.Core;
using iTextSharp.text.pdf;
using System;
using System.IO;

namespace Excel2PDF.PDFProcessor
{
    public class PDFProcessor : IPDFProcessor
    {
        public void GenerateIntroPDF(Excel2PDFConfig config, InfoTemp infoTemp, IntroAcroFields introTempInfo, string newFileName)
        {
            string introTemplate = $"{config.PDFTemplateFolder}\\{introTempInfo.SheetName}";
            string pdfOutPath = $"{config.PDFWriteFolder}\\{newFileName}";
            PdfReader pdfReader = new PdfReader(introTemplate);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(pdfOutPath, FileMode.Create));
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            var date = DateTime.Parse(infoTemp.Date);

            pdfFormFields.SetField(introTempInfo.ProposedForAcroField1, infoTemp.ProposalFor);
            pdfFormFields.SetField(introTempInfo.ProposedForAcroField2, infoTemp.ProposalFor);

            pdfFormFields.SetField(introTempInfo.ProposedByAcroField, infoTemp.ProposalBy);
            pdfFormFields.SetField(introTempInfo.DateAcroField, date.ToString(config.DateFormat));

            pdfStamper.FormFlattening = false;
            pdfStamper.Close();
        }
    }
}
