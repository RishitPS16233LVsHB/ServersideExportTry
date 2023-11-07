using System;
using OfficeOpenXml;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;


namespace ServerSideExportTry
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            
            DataTable dataTable = new DataTable("sample Data table");
            dataTable.Columns.Add("id", typeof(int));
            dataTable.Columns.Add("name", typeof(string));
            dataTable.Columns.Add("age", typeof(int));
            dataTable.Columns.Add("salary", typeof(int));


            dataTable.Rows.Add(new object[] { 1, "rishit", 21, 14700});
            dataTable.Rows.Add(new object[] { 2, "hetvi", 21, 40000 });
            dataTable.Rows.Add(new object[] { 3, "thanuja", 21, 40000 });
            dataTable.Rows.Add(new object[] { 3, "raj", 21, 20000});
            dataTable.Rows.Add(new object[] { 3, "muneer", 21, 25000});

            using (ExcelPackage pack = new ExcelPackage())
            { 
                ExcelWorksheet worksheet = pack.Workbook.Worksheets.Add("employee details");
                worksheet.Cells["A1"].LoadFromDataTable(dataTable);

                FileInfo fileInfo = new FileInfo("Employee.xlsx");
                pack.SaveAs(fileInfo);



            }

            string htmlContent = "<html><head><title>Sample HTML to PDF</title></head><body><h1>Hello, World!</h1><img src=\"D:\\ServerSideExportTry\\ServerSideExportTry\\bin\\Debug\\net5.0\\Image\\Geralt.jpg\"/></body></html>";
            string outputPath = "output.pdf";

            GeneratePdfFromHtml(htmlContent, outputPath);
            Console.WriteLine("Excel and PDF generated successfully.");
            Console.Read();
        }

        public static void GeneratePdfFromHtml(string htmlContent, string outputPath)
        {
            using (FileStream fs = new FileStream(outputPath, FileMode.Create))
            {
                using (Document doc = new Document())
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    using (StringReader sr = new StringReader(htmlContent))
                    {
                        XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, sr);
                    }

                    doc.Close();
                }
            }
        }
    }
}
