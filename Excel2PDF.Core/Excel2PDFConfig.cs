using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2PDF.Core
{
    public class Excel2PDFConfig
    {
        public string ExcelReadFolder { get; set; }

        public string ExcelArchiveFolder { get; set; }

        public string PDFTemplateFolder { get; set; }

        public string PDFWriteFolder { get; set; }

        public string DateFormat { get; set; }
    }
}
