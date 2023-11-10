namespace CSVProgram
{
    using System;

    public class Employee
    {
        // Unique identifier for the employee
        public int Id { get; set; }

        // Full name of the employee
        public string Name { get; set; }

        // Job title or role of the employee
        public string Role { get; set; }

        // Annual salary of the employee
        public decimal Salary { get; set; }

        // Department in which the employee works
        public string Department { get; set; }

        // Work email address of the employee
        public string Email { get; set; }

        // Constructor for creating an employee
        public Employee(int id, string name, string role, decimal salary, string department, string email)
        {
            Id = id;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Role = role ?? throw new ArgumentNullException(nameof(role));
            Department = department ?? throw new ArgumentNullException(nameof(department));
            Email = email ?? throw new ArgumentNullException(nameof(email));
        }
    }

}