using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeOpenXml;
namespace CSVProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            // List<Employee> employees = new List<Employee>
            // {
            //     new Employee(1, "John Doe", "Software Developer", 55000, "Engineering", "john.doe@example.com"),
            //     new Employee(2, "Jane Smith", "Project Manager", 65000, "Marketing", "jane.smith@example.com"),
            //     new Employee(3, "Mike Brown", "Quality Assurance", 60000, "Engineering", "mike.brown@example.com"),
            //     new Employee(4, "Sarah Miller", "UI/UX Designer", 57000, "Design", "sarah.miller@example.com"),
            //     new Employee(5, "Olivia Jones", "Data Analyst", 62000, "Data Science", "olivia.jones@example.com"),
            //     new Employee(6, "Daniel Garcia", "Sales Representative", 48000, "Sales", "daniel.garcia@example.com"),
            //     new Employee(7, "Emma Wilson", "HR Specialist", 53000, "Human Resources", "emma.wilson@example.com"),
            //     new Employee(8, "Carlos Sanchez", "Customer Support", 45000, "Support", "carlos.sanchez@example.com"),
            //     new Employee(9, "Sophia Davis", "Marketing Coordinator", 50000, "Marketing", "sophia.davis@example.com"),
            //     new Employee(10, "James Hall", "Chief Technology Officer", 95000, "Executive", "james.hall@example.com"),
            // };

            var filePath = @"..\..\..\my_employees.xlsx";
            var excelWriter = new ExcelWriter();
            
            // Read employees from the excel
            List<Employee> employees = GetEmployees();
            DisplayEmployees(employees);

            Console.WriteLine("Enter the ID of the employee you want to update:");
            int employeeId = Convert.ToInt32(Console.ReadLine());

            Employee selectedEmployee = employees.Find(e => e.Id == employeeId);
            if (selectedEmployee != null)
            {
                Console.WriteLine("Enter the new name (leave blank to keep existing):");
                string newName = Console.ReadLine();
                selectedEmployee.Name = string.IsNullOrWhiteSpace(newName) ? selectedEmployee.Name: newName;

                Console.WriteLine("Enter the new role (leave blank to keep existing):");
                string newRole = Console.ReadLine();
                selectedEmployee.Role = string.IsNullOrWhiteSpace(newRole) ? selectedEmployee.Role: newRole;

                Console.WriteLine("Enter the new salary (leave blank to keep existing):");
                string newSalStr = Console.ReadLine();

                if (!string.IsNullOrWhiteSpace(newSalStr))
                {
                    if (decimal.TryParse(newSalStr, out decimal newSal))
                    {
                        selectedEmployee.Salary = newSal;
                    }
                    else
                    {
                        Console.WriteLine("Invalid salary format. Keeping existing salary.");
                    }
                }

                Console.WriteLine("Enter the new department (leave blank to keep existing):");
                string newDep = Console.ReadLine();
                selectedEmployee.Department = string.IsNullOrWhiteSpace(newDep) ? selectedEmployee.Department: newDep;

                Console.WriteLine("Enter the new email (leave blank to keep existing):");
                string newEmail = Console.ReadLine();
                selectedEmployee.Email = string.IsNullOrWhiteSpace(newEmail) ? selectedEmployee.Email: newEmail;

                // Update the employee
                excelWriter.UpdateEmployeeInExcel(filePath, employeeId, selectedEmployee);
                Console.WriteLine("Employee updated successfully!");
            }
            else
            {
                Console.WriteLine("Employee not found.");
            }

            // excelWriter.WriteEmployeesToExcel(filePath, employees);
            List<Employee> newEmployees = GetEmployees();
            DisplayEmployees(newEmployees);
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
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
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

                        
                        // Console.WriteLine($"ID: {employee.Id}, Name: {employee.Name}, Role: {employee.Role}, " +
                        //                 $"Salary: {employee.Salary}, Department: {employee.Department}, " +
                        //                 $"Email: {employee.Email}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while reading in the excel file: {ex.Message}");
            }

            return employees;
        }

        // Display Employees
        private static void DisplayEmployees(List<Employee> employees)
        {
            foreach (var employee in employees)
            {
                Console.WriteLine($"ID: {employee.Id}, Name: {employee.Name}, Role: {employee.Role}, " +
                                        $"Salary: {employee.Salary}, Department: {employee.Department}, " +
                                        $"Email: {employee.Email}");
            }
        }
    }
}