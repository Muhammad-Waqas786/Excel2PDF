# Excel2PDF

The project takes in excel files (from a folder: **ExcelReadFolder**), extracts data and then fills a PDF template (from folder: **PDFTemplateFolder**) with the extracted data. The output PDF files are stored in folder: **PDFWriteFolder**. The read files are later moved to location: **ExcelArchiveFolder**. The solution contains the following sub projects:

| Project | Type |
|:--------| :-------------|
| Excel2PDF | Console Application
| Excel2PDF.Core | Class Library
| Excel2PDF.ExcelProcessor | Class Library
| Excel2PDF.PDFProcessor | Class Library

This application can handle both xls and xlsx files. The following libraries are used to process Excel and PDF files

| Nuget Package | Version | Downloads |
|:--------| :-------------| :-------------|
| iTextSharp | 5.5.13.2 | 16,345,362
| ExcelDataReader | 3.7.0-develop00310 | 13,636,130
| ExcelDataReader.DataSet | 3.6.0 | 7,798,405