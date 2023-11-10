using ClosedXML.Excel;

namespace CSVProgram
{
    public class ExcelWriter
    {
        public void WriteEmployeesToExcel(string filePath, List<Employee> employees)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Employees");

                // Define headers
                worksheet.Cell("A1").Value = "Id";
                worksheet.Cell("B1").Value = "Name";
                worksheet.Cell("C1").Value = "Role";
                worksheet.Cell("D1").Value = "Salary";
                worksheet.Cell("E1").Value = "Department";
                worksheet.Cell("F1").Value = "Email";

                int currentRow = 2;
                foreach (var employee in employees)
                {
                    worksheet.Cell(currentRow, 1).Value = employee.Id;
                    worksheet.Cell(currentRow, 2).Value = employee.Name;
                    worksheet.Cell(currentRow, 3).Value = employee.Role;
                    worksheet.Cell(currentRow, 4).Value = employee.Salary;
                    worksheet.Cell(currentRow, 5).Value = employee.Department;
                    worksheet.Cell(currentRow, 6).Value = employee.Email;

                    currentRow++;
                }

                workbook.SaveAs(filePath);
            }
        }
    }
}
