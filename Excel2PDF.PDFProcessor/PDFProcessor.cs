using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using Excel2PDF.Core;
using System.IO;
using System.Globalization;

namespace Excel2PDF.PDFProcessor
{
    public class PDFProcessor : IPDFProcessor
    {
        private readonly string introTemplateName = "Intro.pdf";
        private readonly string proposedForAcroField1 = "Text1";
        private readonly string proposedForAcroField2 = "Text1a";
        private readonly string proposedByAcroField = "Text2";
        private readonly string dateAcroField = "Text3";

        public void GenerateIntroPDF(Excel2PDFConfig config, InfoTemp infoTemp, string newFileName)
        {
            string introTemplate = $"{config.PDFTemplateFolder}\\{introTemplateName}";
            string pdfOutPath = $"{config.PDFWriteFolder}\\{newFileName}";
            PdfReader pdfReader = new PdfReader(introTemplate);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(pdfOutPath, FileMode.Create));
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            var date = DateTime.Parse(infoTemp.Date);

            pdfFormFields.SetField(proposedForAcroField1, infoTemp.ProposalFor);
            pdfFormFields.SetField(proposedForAcroField2, infoTemp.ProposalFor);

            pdfFormFields.SetField(proposedByAcroField, infoTemp.ProposalBy);
            pdfFormFields.SetField(dateAcroField, date.ToString(config.DateFormat));

            pdfStamper.FormFlattening = false;
            pdfStamper.Close();
        }
    }
}
