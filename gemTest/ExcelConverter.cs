using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using System.IO;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace gemTest
{
    class ExcelConverter
    {
        public string filePathXls { get; set; }
        public string filePathXlsX { get; set; }
        // Consts that deal with column locations in the excel file given by GSO

        const int activeCol = 4;
        const int emailNameCol = 2;
        const int emailCol = 3;
        const int destCol = 22;
        
        public ExcelConverter(string filePathXls, string filePathXlsX)
        {
            this.filePathXls = filePathXls;
            this.filePathXlsX = filePathXlsX;
        }
        public ExcelConverter(string filePathXlsX)
        {
            this.filePathXlsX = filePathXlsX;
        }
        public void convertDocument()
        {
            var excelApplication = new Excel.Application();
            excelApplication.Visible = false;
            excelApplication.DisplayAlerts = false;
            var wbk = excelApplication.Workbooks.Open(filePathXls); 
           
            wbk.SaveAs(filePathXlsX, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wbk.Close();
            excelApplication.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(wbk);
            Marshal.ReleaseComObject(excelApplication);
            // Prevent memory leaks
        }

        public void consolidateEmails(string activeFilePathXlsX, string emailFilePathXlsX)
        {
            FileInfo filePathA = new FileInfo(activeFilePathXlsX);
            FileInfo filePathB = new FileInfo(emailFilePathXlsX);

            try
            {
                // uses var to prevent memory leaks http://stackoverflow.com/questions/13483523/c-sharp-excel-automation-causes-excel-memory-leak
                var activeFilePath = new ExcelPackage(filePathA);
                var emailFilePath = new ExcelPackage(filePathB);
                var activeWorksheet = activeFilePath.Workbook.Worksheets[1];
                var emailWorkSheet = emailFilePath.Workbook.Worksheets[1];


                var activeStart = activeWorksheet.Dimension.Start;
                var activeEnd = activeWorksheet.Dimension.End;
                var emailStart = emailWorkSheet.Dimension.Start;
                var emailEnd = emailWorkSheet.Dimension.End;

               // Console.WriteLine(worksheet.Cells[activeRow, activeCol].Text);
                for (int activeRow = activeStart.Row; activeRow <= activeEnd.Row; activeRow++)
                { // Row by row...
                    for (int emailRow = emailStart.Row; emailRow <= emailEnd.Row; emailRow++)
                    {
                        if (activeWorksheet.Cells[activeRow, activeCol].Text.Contains(emailWorkSheet.Cells[emailRow, emailNameCol].Text) || emailWorkSheet.Cells[emailRow, emailNameCol].Text.Contains(activeWorksheet.Cells[activeRow, activeCol].Text))
                        {
                            activeWorksheet.Cells[activeRow, destCol].Value = emailWorkSheet.Cells[emailRow, emailCol].Text;
                        }
                    }        
                }
                activeFilePath.Save();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

    }
}
