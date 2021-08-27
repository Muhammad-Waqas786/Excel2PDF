using Excel2PDF.Core;
using System.IO;

namespace Excel2PDF.ExcelProcessor
{
    public interface IExcelProcessor
    {
        InfoTemp ParseExcelForInfo(FileStream stream);        
    }
}
