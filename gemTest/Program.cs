using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using System.Data.SqlClient;
using System.Data;
using System.Web;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Globalization;
using System.Threading;

namespace gemTest
{
    class Program
    {
        static void Main(string[] args)
        {

            string smtp = "bogus";
            string activeFilePathXls = @"C:\bogus\activityreport.xls";
            string activeFilePathXlsX = @"C:\bogus\activityreport.xlsx";
            string emailFilePathXls = @"C:\bogus\ActEmails.xls";
            string emailFilePathXlsX = @"C:\bogus\ActEmails.xlsx";
           

            string fileName = "client.htm";
            string path = Path.Combine(Directory.GetCurrentDirectory(), @"Data\", fileName);

            Outlook.Application application = new Outlook.Application();
            SendMail sendy = new SendMail(application, "test", "test", "alex@tjhpayroll.com", smtp);
            CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;
            TextInfo textInfo = cultureInfo.TextInfo;
            List<EmailList> lstemail = new List<EmailList>();

            ExcelConverter excelConverterActiveReport = new ExcelConverter(activeFilePathXls, activeFilePathXlsX);
            ExcelConverter excelConverterEmailList = new ExcelConverter(emailFilePathXls, emailFilePathXlsX);
            ExcelExtractor extractCompInfo = new ExcelExtractor(activeFilePathXlsX);
            try
            {
                excelConverterActiveReport.convertDocument();
                excelConverterEmailList.convertDocument();
                excelConverterActiveReport.consolidateEmails(excelConverterActiveReport.filePathXlsX, excelConverterEmailList.filePathXlsX);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            try
            {
                // loads the converted XLSX GSO report and dumps the company info into a list of my EmailList class.
                lstemail = extractCompInfo.extractCompanyInfo(lstemail);
                foreach (var item in lstemail)
                {
                    Console.WriteLine("Company name is {0} and the tracking number is {1}", item.compName, item.trackNumber);
                }

                Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            try
            {
                int emailCount = 0;
                foreach (var item in lstemail)
                {
                    if (item.compEmail != null)
                    {
                        string Body = System.IO.File.ReadAllText(path);
                        var itemTestName = textInfo.ToLower(item.compName);
                        var itemTestNumber = item.trackNumber;
                        string itemTestContName;
                        string properName = null;
                        if (item.contactName != null)
                        {
                            itemTestContName = textInfo.ToLower(item.contactName);
                            properName = textInfo.ToTitleCase(itemTestContName);
                        }
                        var properCompName = textInfo.ToTitleCase(itemTestName);

                        Body = Body.Replace("#DealerCompanyName#", itemTestName);
                        Body = Body.Replace("#DealerTrackingNumber#", itemTestNumber);
                        if (properName == null || properName == String.Empty)
                            Body = Body.Replace("#DealerName#", properCompName);
                        else
                            Body = Body.Replace("#DealerName#", properName);
                        Body = Body.Replace("#TodayDate#", DateTime.Now.ToString());
                        sendy.body = Body;
                        sendy.to = item.compEmail;  

                        sendy.subject = "Your payroll is out for delivery";
                        sendy.SendEmailFromAccount();
                        emailCount++;
                        Console.WriteLine("Sent email number " + emailCount);
                        //break; 
                    }
                }
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
