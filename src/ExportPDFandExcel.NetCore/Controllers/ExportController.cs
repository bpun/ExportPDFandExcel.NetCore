using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;


// For more information on enabling MVC for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace ExportPDFandExcel.NetCore.Controllers
{
    public class ExportController : Controller
    {
        // GET: /<controller>/
        public IActionResult Index()
        {
            var dataList = GetStudentList();
            return View(dataList);
        }


        public FileContentResult ExportExcel()
        {
            //student list
            var dataList = GetStudentList();

            //column Header name
            var columnsHeader = new List<string>{
                "S/N",
                "Name",
                "Address",
                "Phone",
                "Grade"
            };
            var filecontent = ExportExcel(dataList, columnsHeader, "Students");
            return File(filecontent, "application/ms-excel", "students.xlsx"); ;
        }

        public FileContentResult ExportPDF()
        {
            //student list
            var dataList = GetStudentList();

            //column Header name
            var columnsHeader = new List<string>{
                "S/N",
                "Name",
                "Address",
                "Phone",
                "Grade"
            };
            var filecontent = ExportPDF(dataList, columnsHeader, "Students");
            return File(filecontent, "application/pdf", "students.pdf"); ;
        }



        #region Helpers 
        private List<StudentModel> GetStudentList()
        {
            var studentList = new List<StudentModel>();
            var student1 = new StudentModel
            {
                Id = 1,
                Name = "Ram",
                Address = "Ktm",
                Phone = "93994948928",
                Grade = "A"
            };
            var student2 = new StudentModel
            {
                Id = 1,
                Name = "Ram",
                Address = "Ktm",
                Phone = "93994948928",
                Grade = "A"
            };
            var student3 = new StudentModel
            {
                Id = 1,
                Name = "Ram",
                Address = "Ktm",
                Phone = "93994948928",
                Grade = "A"
            };
            var student4 = new StudentModel
            {
                Id = 1,
                Name = "Ram",
                Address = "Ktm",
                Phone = "93994948928",
                Grade = "A"
            };
            studentList.Add(student1);
            studentList.Add(student2);
            studentList.Add(student3);
            studentList.Add(student4);


            return studentList;
        }

        private byte[] ExportExcel(List<StudentModel> dataList, List<string> columnsHeader, string heading)
        {
            byte[] result = null;

            using (ExcelPackage package = new ExcelPackage())
            {
                // add a new worksheet to the empty workbook
                var worksheet = package.Workbook.Worksheets.Add(heading);
                using (var cells = worksheet.Cells[1, 1, 1, 5])
                {
                    cells.Style.Font.Bold = true;
                    cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cells.Style.Fill.BackgroundColor.SetColor(Color.Green);
                }
                //First add the headers
                for (int i = 0; i < dataList.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataList[i];
                }

                //Add values
                var j = 2;
                var count = 1;
                foreach (var item in dataList)
                {
                    worksheet.Cells["A" + j].Value = count;
                    worksheet.Cells["B" + j].Value = item.Name;
                    worksheet.Cells["C" + j].Value = item.Address;
                    worksheet.Cells["D" + j].Value = item.Phone;
                    worksheet.Cells["E" + j].Value = item.Grade;

                    j++;
                    count++;
                }
                result = package.GetAsByteArray();
            }

            return result;
        }
        public class StudentModel
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Address { get; set; }
            public string Phone { get; set; }
            public string Grade { get; set; }

        }


        private byte[] ExportPDF(List<StudentModel> dataList, List<string> columnsHeader, string heading)
        {

            var document = new Document();
            var outputMS = new MemoryStream();
            var writer = PdfWriter.GetInstance(document, outputMS);
            document.Open();
            var font5 = FontFactory.GetFont(FontFactory.HELVETICA, 11);

            document.Add(new Phrase(Environment.NewLine));

            //var count = typeof(UserListVM).GetProperties().Count();
            var count = columnsHeader.Count;
            var table = new PdfPTable(count);
            float[] widths = new float[] { 2f, 4f, 5f, 4f, 4f };

            table.SetWidths(widths);

            table.WidthPercentage = 100;
            var cell = new PdfPCell(new Phrase(heading));
            cell.Colspan = count;

            for (int i = 0; i < count; i++)
            {
                var headerCell = new PdfPCell(new Phrase(columnsHeader[i], font5));
                headerCell.BackgroundColor = BaseColor.Gray;
                table.AddCell(headerCell);
            }

            var sn = 1;
            foreach (var item in dataList)
            {
                table.AddCell(new Phrase(sn.ToString(), font5));
                table.AddCell(new Phrase(item.Name, font5));
                table.AddCell(new Phrase(item.Address, font5));
                table.AddCell(new Phrase(item.Phone, font5));
                table.AddCell(new Phrase(item.Grade, font5));

                sn++;
            }

            document.Add(table);
            document.Close();
            var result = outputMS.ToArray();

            return result;
        }

        #endregion
    }
}
