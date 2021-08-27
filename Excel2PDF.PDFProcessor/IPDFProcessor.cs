using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel2PDF.Core;

namespace Excel2PDF.PDFProcessor
{
    public interface IPDFProcessor
    {
        void GenerateIntroPDF(Excel2PDFConfig config, InfoTemp infoTemp, string newFileName);
    }
}
