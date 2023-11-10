using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CSVProgram
{
    public class ExcelReader
    {
        public List<Employee> ReadEmployeesFromExcel(string filePath)
        {
            var employees = new List<Employee>();

            var fileInfo = new FileInfo(filePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                var workbook = package.Workbook;
                if (workbook.Worksheets.Count == 0)
                {
                    throw new Exception("No worksheet found.");
                }

                var worksheet = workbook.Worksheets[0];

                // Assuming the first row is the header and data starts from the second row
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    try
                    {
                        var id = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        var name = worksheet.Cells[row, 2].Value.ToString();
                        var role = worksheet.Cells[row, 3].Value.ToString();
                        var salary = Convert.ToDecimal(worksheet.Cells[row, 4].Value, CultureInfo.InvariantCulture);
                        var department = worksheet.Cells[row, 5].Value.ToString();
                        var email = worksheet.Cells[row, 6].Value.ToString();

                        employees.Add(new Employee(id, name, role, salary, department, email));
                    }
                    catch (FormatException fe)
                    {
                        Console.WriteLine($"Data format error in row {row}: {fe.Message}");
                    }
                    catch (InvalidCastException ice)
                    {
                        Console.WriteLine($"Invalid cast in row {row}: {ice.Message}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unexpected error in row {row}: {ex.Message}");
                    }
                }
            }

            return employees;
        }
    }
}
