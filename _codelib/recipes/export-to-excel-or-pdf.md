# Export data to Excel, report to PDF

Keywords: Excel export, recordset to Excel, CopyFromRecordset, Excel template, PDF export, report as PDF, OutputTo,
nach Excel exportieren, Datensätze nach Excel, Excel-Vorlage, Bericht als PDF, PDF erzeugen

## Import

- `api/office/excel/modExcel_Tools.bas` (Excel export)
- for PDF usually nothing (built-in Access export)

## Usage

```vba
' Recordset (DAO or ADODB) to a new Excel workbook (Excel stays open for the user)
Dim rst As DAO.Recordset
Set rst = CurrentDb.OpenRecordset("SELECT * FROM qryOrderExport", dbOpenSnapshot)
RecordsetExcelExport rst
rst.Close

' ... into a template, starting at row 3, without header row, into a named sheet
RecordsetExcelExport rst, CurrentProject.Path & "\templates\Orders.xltx", 3, 1, False, "Data"

' Report as PDF with Access built-in export (Access 2007 SP2 or later)
DoCmd.OpenReport "rptInvoice", acViewPreview, , "InvoiceID = 4711", acHidden
DoCmd.OutputTo acOutputReport, "rptInvoice", acFormatPDF, "C:\Temp\Invoice_4711.pdf"
DoCmd.Close acReport, "rptInvoice"
```

## Notes

- `RecordsetExcelExport(Recordset, TemplateFile, StartRow, StartCol, WithRecordsetHeaders, ImportTabName)` uses late
  binding: no reference to the Excel library required (compiler constant `ExcelEarlyBinding = 1` for early binding).
- Build the filter of the report with SqlTools (see [sql-criteria.md](sql-criteria.md)).
- The PdfCreator package (`api/pdfcreator/LibPdfCreatorTools.cls`) uses the COM interface of PDFCreator 1.x and is only
  needed for special requirements (other formats, PDFCreator profiles).

## See also

- `api/pdfcreator/LibPdfCreatorTools.cls`, `_examples/api/pdfcreator/PdfCreatorExample.frm`
