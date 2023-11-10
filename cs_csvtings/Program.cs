using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeOpenXml;
namespace CSVProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Employee> employees = new List<Employee>
            {
                new Employee(1, "John Doe", "Software Developer", 55000, "Engineering", "john.doe@example.com"),
                new Employee(2, "Jane Smith", "Project Manager", 65000, "Marketing", "jane.smith@example.com"),
                new Employee(3, "Mike Brown", "Quality Assurance", 60000, "Engineering", "mike.brown@example.com"),
                new Employee(4, "Sarah Miller", "UI/UX Designer", 57000, "Design", "sarah.miller@example.com"),
                new Employee(5, "Olivia Jones", "Data Analyst", 62000, "Data Science", "olivia.jones@example.com"),
                new Employee(6, "Daniel Garcia", "Sales Representative", 48000, "Sales", "daniel.garcia@example.com"),
                new Employee(7, "Emma Wilson", "HR Specialist", 53000, "Human Resources", "emma.wilson@example.com"),
                new Employee(8, "Carlos Sanchez", "Customer Support", 45000, "Support", "carlos.sanchez@example.com"),
                new Employee(9, "Sophia Davis", "Marketing Coordinator", 50000, "Marketing", "sophia.davis@example.com"),
                new Employee(10, "James Hall", "Chief Technology Officer", 95000, "Executive", "james.hall@example.com"),
            };

            var filePath = @"..\..\..\my_employees.xlsx";
            var excelWriter = new ExcelWriter();
            

            excelWriter.WriteEmployeesToExcel(filePath, employees);
            GetEmployees();
        }

        // Method to get employees
        public static List<Employee> GetEmployees()
        {
            var employees = new List<Employee>();
            var filePath = @"..\..\..\my_employees.xlsx";

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Assuming first row is header
                    {
                        Employee employee = new(
                            id: Convert.ToInt32(worksheet.Cells[row, 1].Value?.ToString()),
                            name: worksheet.Cells[row, 2].Value?.ToString(),
                            role: worksheet.Cells[row, 3].Value?.ToString(),
                            salary: Convert.ToDecimal(worksheet.Cells[row, 4].Value?.ToString()),
                            department: worksheet.Cells[row, 5].Value?.ToString(),
                            email: worksheet.Cells[row, 6].Value?.ToString()
                        );
                        employees.Add(employee);

                        // Optional: Print out each employee's details
                        Console.WriteLine($"ID: {employee.Id}, Name: {employee.Name}, Role: {employee.Role}, " +
                                        $"Salary: {employee.Salary}, Department: {employee.Department}, " +
                                        $"Email: {employee.Email}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while reading in the excel file: {ex.Message}");
            }

            return employees;
        }
    }
}