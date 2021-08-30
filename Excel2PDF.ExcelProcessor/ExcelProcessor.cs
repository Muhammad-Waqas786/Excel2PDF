using Excel2PDF.Core;
using ExcelDataReader;
using System;
using System.IO;

namespace Excel2PDF.ExcelProcessor
{
    public class ExcelProcessor : IExcelProcessor
    {
        public InfoTemp ParseExcelForInfo(FileStream stream)
        {
            InfoTemp infoTemp = new InfoTemp();
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                var dataTables = result.Tables;
                var infoTable = dataTables["Info"];
                infoTemp.ProposalFor = Convert.ToString(infoTable.Rows[1][1]);
                infoTemp.ProposalBy = Convert.ToString(infoTable.Rows[2][1]);
                infoTemp.Date = Convert.ToString(infoTable.Rows[3][1]);
            }

            return infoTemp;
        }
    }
}
