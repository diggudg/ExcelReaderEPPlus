using ExcelReaderUsingOpenOfficeXML.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace RefreshOauth.Pluggins.Excel
{
    public class ExcelServices
    {
        public List<Student> ReadUploadedExcel(HttpPostedFile file)
        {
            var usersList = new List<Student>();
            using (var package = new ExcelPackage(file.InputStream))
            {
                var currentSheet = package.Workbook.Worksheets;
                var workSheet = currentSheet.First();
                var noOfCol = workSheet.Dimension.End.Column;
                var noOfRow = workSheet.Dimension.End.Row;
                //json = new JavaScriptSerializer().Serialize(workSheet);
                for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                {
                    var user = new Student();
                    // user.ClientId =Convert.ToInt32( workSheet.Cells[rowIterator, 1].Value.ToString());
                    user.ClientId = workSheet.Cells[rowIterator, 1].Value != null ? Convert.ToInt32(workSheet.Cells[rowIterator, 1].Value.ToString()) : 0;
                    user.StudName = workSheet.Cells[rowIterator, 2].Value != null ? workSheet.Cells[rowIterator, 2].Value.ToString() : "";
                    user.StudMiddleName = workSheet.Cells[rowIterator, 3].Value != null ? workSheet.Cells[rowIterator, 3].Value.ToString() : "";
                    user.StudLastName = workSheet.Cells[rowIterator, 4].Value != null ? workSheet.Cells[rowIterator, 4].Value.ToString() : "";
                    user.StudFathername = workSheet.Cells[rowIterator, 5].Value != null ? workSheet.Cells[rowIterator, 5].Value.ToString() : "";
                    user.StudMotherName = workSheet.Cells[rowIterator, 6].Value != null ? workSheet.Cells[rowIterator, 6].Value.ToString() : "";
                    user.DOB = workSheet.Cells[rowIterator, 7].Value != null ? Convert.ToDateTime(workSheet.Cells[rowIterator, 7].Value.ToString()) : Convert.ToDateTime("0000-00-00");
                    usersList.Add(user);
                }
            }
            return usersList;
        }

        public  DataTable ToDataTable(HttpPostedFile file)
        {
            var package = new ExcelPackage(file.InputStream);
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable table = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }
            for (int i = 1; i <= workSheet.Dimension.End.Column; i++)
            {
                table.Columns.Add(i.ToString());
            }
            for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }
            return table;
        }

        public List<StudentsDetails> ReadUploadedExcelTest(HttpPostedFile file)
        {
            var usersList = new List<StudentsDetails>();
            using (var package = new ExcelPackage(file.InputStream))
            {
                var currentSheet = package.Workbook.Worksheets;
                var workSheet = currentSheet.First();
                var noOfCol = workSheet.Dimension.End.Column;
                var noOfRow = workSheet.Dimension.End.Row;
                //json = new JavaScriptSerializer().Serialize(workSheet);
                for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                {
                    var user = new StudentsDetails();
                    // user.ClientId =Convert.ToInt32( workSheet.Cells[rowIterator, 1].Value.ToString());
                    user.Clients.ClientID = workSheet.Cells[rowIterator, 1].Value != null ? Convert.ToInt32(workSheet.Cells[rowIterator, 1].Value.ToString()) : 0;
                    //students persoal details
                    user.StudPersonalDetails.Name = workSheet.Cells[rowIterator, 2].Value != null ? workSheet.Cells[rowIterator, 2].Value.ToString() : "";
                    user.StudPersonalDetails.MiddleName = workSheet.Cells[rowIterator, 3].Value != null ? workSheet.Cells[rowIterator, 3].Value.ToString() : "";
                    user.StudPersonalDetails.LastName = workSheet.Cells[rowIterator, 4].Value != null ? workSheet.Cells[rowIterator, 4].Value.ToString() : "";
                    user.StudPersonalDetails.FatherName= workSheet.Cells[rowIterator, 5].Value != null ? workSheet.Cells[rowIterator, 5].Value.ToString() : "";
                    user.StudPersonalDetails.MotherName = workSheet.Cells[rowIterator, 6].Value != null ? workSheet.Cells[rowIterator, 6].Value.ToString() : "";
                    user.StudPersonalDetails.DOB = workSheet.Cells[rowIterator, 7].Value != null ? workSheet.Cells[rowIterator, 7].Value.ToString() : "0000-00-00";
                    //students Contacts
                    user.StudentContacts.Phone = workSheet.Cells[rowIterator, 8].Value != null ? workSheet.Cells[rowIterator, 8].Value.ToString() : ""; ;
                    user.StudentContacts.AlternatePhone = workSheet.Cells[rowIterator, 9].Value != null ? workSheet.Cells[rowIterator, 9].Value.ToString() : ""; ;
                    user.StudentContacts.Email = workSheet.Cells[rowIterator, 10].Value != null ? workSheet.Cells[rowIterator, 10].Value.ToString() : ""; ;
                    //Students Address
                    user.StudentAddress.Address1 = workSheet.Cells[rowIterator, 11].Value != null ? workSheet.Cells[rowIterator, 11].Value.ToString() : ""; ;
                    user.StudentAddress.Address2 = workSheet.Cells[rowIterator, 12].Value != null ? workSheet.Cells[rowIterator, 12].Value.ToString() : ""; ;
                    user.StudentAddress.City = workSheet.Cells[rowIterator, 13].Value != null ? workSheet.Cells[rowIterator, 13].Value.ToString() : ""; ;
                    user.StudentAddress.State = workSheet.Cells[rowIterator, 14].Value != null ? workSheet.Cells[rowIterator, 14].Value.ToString() : ""; ;
                    user.StudentAddress.PIN = workSheet.Cells[rowIterator, 15].Value != null ?Convert.ToInt32( workSheet.Cells[rowIterator, 15].Value.ToString()) : 000000; ; 
                    
                    // course Details
                    user.CourseDetails.CourseID = workSheet.Cells[rowIterator, 16].Value != null ?Convert.ToInt32( workSheet.Cells[rowIterator, 16].Value.ToString()):00; ;
                    user.CourseDetails.CourseName = workSheet.Cells[rowIterator, 17].Value != null ? workSheet.Cells[rowIterator, 17].Value.ToString() : ""; ;



                    usersList.Add(user);
                }
            }
            return usersList;
        }

    }
    public static class ExcelPackageExtensions
    {
        public static DataTable ToDataTable(this ExcelPackage package)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable table = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }

            for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }
            return table;
        }
    }
}