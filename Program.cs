using System;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System.Data.Common;
using GemBox.Document;



namespace ServerSideExportTry
{
    internal class Program
    {
        static void Main(string[] args)
        {

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            DocumentModel.Load("Hello_world.docx").Save("output1.pdf");
            return;
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

            //string query = "select * from Sales.order_items"
            using (SqlConnection connection = new SqlConnection("Data Source=DESKTOP-T8Q9USL;Initial Catalog=BikeStoreDatabase;Integrated Security=True"))
            {
                dataTable.Rows.Clear();
                dataTable.Columns.Clear();
                dataTable.Clear();
                SqlDataAdapter adp = new SqlDataAdapter("select * from Sales.order_items", connection);
                adp.Fill(dataTable);                
            }

            // this is license based but we can use NonCommercial one to solve our problem
            // non commercial option is based on 5th version of EPPlus and provides unlimited use to Excel export and many styling options
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage pack = new ExcelPackage())
            { 
                ExcelWorksheet worksheet = pack.Workbook.Worksheets.Add("employee details");
                System.Collections.Generic.List<string> columnName = new System.Collections.Generic.List<string>();
                foreach (DataColumn dc in dataTable.Columns)
                    columnName.Add(dc.ColumnName);

                worksheet.Cells["A1"].LoadFromCollection<string>(columnName);
                worksheet.Cells["A2"].LoadFromDataTable(dataTable);

                FileInfo fileInfo = new FileInfo("Employee.xlsx");
                pack.SaveAs(fileInfo);
            }


            string htmlContent = "<html><head><title>Sample HTML to PDF</title></head><body><h1>Hello, World!</h1><img src=\"D:\\ServerSideExportTry\\ServerSideExportTry\\bin\\Debug\\net5.0\\Image\\Geralt.jpg\"/></body></html>";
            string outputPath = "output.pdf";

            GeneratePdfFromDataTable(dataTable, outputPath);





            Console.WriteLine("Excel and PDF generated successfully.");
            Console.Read();
        }
        public static void GeneratePdfFromDataTable(DataTable dataTable, string outputPath)
        {
            using (FileStream fs = new FileStream(outputPath, FileMode.Create))
            {
                using (Document doc = new Document())
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    // Create a PDF table
                    PdfPTable pdfTable = new PdfPTable(dataTable.Columns.Count);

                    // Add table headers
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(column.ColumnName));
                        pdfTable.AddCell(cell);
                    }

                    // Add table data
                    foreach (DataRow row in dataTable.Rows)
                    {
                        foreach (object item in row.ItemArray)
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(item.ToString()));
                            pdfTable.AddCell(cell);
                        }
                    }

                    // Add the table to the document
                    doc.Add(pdfTable);

                    doc.Close();
                }
            }
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
