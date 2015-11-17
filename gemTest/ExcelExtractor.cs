using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gemTest
{
    class ExcelExtractor
    {

        public string filePathXlsX { get; set; }

        public ExcelExtractor(string filePathXlsX)
        {
            this.filePathXlsX = filePathXlsX;
        }

        public List<EmailList> extractCompanyInfo(List<EmailList> lstemail)
        {
            FileInfo filePathA = new FileInfo(filePathXlsX);

            try
            {
                ExcelPackage activeFilePath = new ExcelPackage(filePathA);
                OfficeOpenXml.ExcelWorksheet activeWorksheet = activeFilePath.Workbook.Worksheets[1];

                var activeStart = activeWorksheet.Dimension.Start;
                var activeEnd = activeWorksheet.Dimension.End;
                for (int activeRow = activeStart.Row; activeRow <= activeEnd.Row; activeRow++)
                { // Row by row    
                    if (activeRow == 1 || activeRow == 2) 
                    {
                        continue;
                    }
                    EmailList data = new EmailList();
                    for (int activeCol = activeStart.Column; activeCol <= activeEnd.Column; activeCol++)
                    { // Cell by cell

                        if (activeWorksheet.Cells[activeRow, activeCol].Text != null && activeCol == 4) // com
                            data.compName = activeWorksheet.Cells[activeRow, activeCol].Text.Trim();
                        else if (activeWorksheet.Cells[activeRow, activeCol].Text != null && activeCol == 1)
                            data.trackNumber = activeWorksheet.Cells[activeRow, activeCol].Text.Trim();
                        else if (activeWorksheet.Cells[activeRow, activeCol].Text != null && activeCol == 9) // contact name
                        {
                            data.contactName = activeWorksheet.Cells[activeRow, activeCol].Text.Trim();
                            // used to get ATTN out of the name
                            if (data.contactName != "" && data.contactName.Length > 6)
                            {
                                if (data.contactName.Substring(0, 5) == "ATTN:")
                                {
                                    data.contactName = data.contactName.Substring(6).Trim();
                                }
                            }
                           
                        }
                        else if (!String.IsNullOrEmpty(activeWorksheet.Cells[activeRow, activeCol].Text) && activeCol == 22)
                            data.compEmail = activeWorksheet.Cells[activeRow, activeCol].Text.Trim();
                    }
                    if (data.trackNumber != null && data.compName != null)
                    {
                        lstemail.Add(data);
                    }
                }               
                return lstemail;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                
                return lstemail;
            }

        }
     
    }
}
