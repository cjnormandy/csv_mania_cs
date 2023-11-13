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

        // Update an Employee
        public void UpdateEmployeeInExcel(string filePath, int employeeId, Employee updatedEmployee)
        {
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet("Employees");

                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
                    bool employeeFound = false;
                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetValue<int>() == employeeId)
                        {
                            row.Cell(2).Value = updatedEmployee.Name;
                            row.Cell(3).Value = updatedEmployee.Role;
                            row.Cell(4).Value = updatedEmployee.Salary;
                            row.Cell(5).Value = updatedEmployee.Department;
                            row.Cell(6).Value = updatedEmployee.Email;
                            employeeFound = true;
                            break;
                        }
                    }
                    
                    if (!employeeFound) 
                    {
                        throw new Exception($"Employee with ID {employeeId} not found.");
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }
        }
    }
}
