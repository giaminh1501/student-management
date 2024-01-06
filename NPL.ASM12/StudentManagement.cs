using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace NPL.ASM12
{
    public class StudentManagement
    {
        public List<Student> students { get; set; }

        public StudentManagement()
        {
            students = new List<Student>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public void Add()
        {
            Student student = new Student();

            do
            {
                Console.Write("Student ID: ");
                student.ID = Console.ReadLine();
                student.CheckID();
            } while (!Regex.IsMatch(student.ID, @"^[a-zA-Z0-9]+$"));

            do
            {
                Console.Write("First name: ");
                student.FirstName = Console.ReadLine();
                student.CheckFirstName();
            } while (!Regex.IsMatch(student.FirstName, @"^[a-zA-Z\s]+$"));

            do
            {
                Console.Write("Last name: ");
                student.LastName = Console.ReadLine();
                student.CheckLastName();
            } while (!Regex.IsMatch(student.LastName, @"^[a-zA-Z\s]+$"));

            do
            {
                Console.Write("Email: ");
                student.Email = Console.ReadLine();
                student.CheckEmail();
            } while (!Regex.IsMatch(student.Email, @"^[a-zA-Z0-9./]+@(outlook\.com|gmail\.com)$"));

            Console.Write("Address: ");
            student.Address = Console.ReadLine();

            do
            {
                Console.Write("Phone number: ");
                student.PhoneNumber = Console.ReadLine();
                student.CheckPhoneNumber();
            } while (!Regex.IsMatch(student.PhoneNumber, @"^(?:\+84|0)[0-9]{9}$"));

            students.Add(student);

            if (students.Count > 0)
            {
                Console.WriteLine("Data added to list successfully. Press any keys to proceed...");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Error: Failed to add data to list. Press any keys to proceed...");
                Console.ReadKey();
            }
        }

        public void Remove()
        {
            if (students.Count > 0)
            {
                Console.Write("Enter a student ID to delete: ");
                string deleteID = Console.ReadLine();

                Student checkID = students.Find(s => s.ID == deleteID);

                // Remove the item if the name is found
                if (checkID != null)
                {
                    students.Remove(checkID);
                    Console.WriteLine($"Student at ID {deleteID} deleted. Press any keys to proceed...");
                    Console.ReadKey();
                }
                else
                {
                    Console.WriteLine("Error: ID not found. Press any keys to proceed...");
                    Console.ReadKey();
                }
            }
            else
            {
                Console.WriteLine("Error: No data in list found. Press any keys to proceed...");
                Console.ReadKey();
            }
        }

        public void Display()
        {
            int count = students.Count();

            if (count == 0)
            {
                Console.WriteLine("Error: No data in list found. Press any keys to proceed...");
                Console.ReadKey();
            }
            else
            {
                int num = 1;
                Console.WriteLine("Students in list:");
                foreach (var student in students)
                {
                    Console.WriteLine($"Record No {num}.");
                    Console.WriteLine("ID: " + student.ID);
                    Console.WriteLine("First name: " + student.FirstName);
                    Console.WriteLine("Last name: " + student.LastName);
                    Console.WriteLine("Email: " + student.Email);
                    Console.WriteLine("Address: " + student.Address);
                    Console.WriteLine("Phone number: " + student.PhoneNumber + "\n");
                    num++;
                }

                Console.WriteLine($"Successfully displayed {count} records in list. Press any keys to proceed...");
                Console.ReadKey();
            }
        }

        public void CheckExit()
        {
            string choice;
            int check = 0;

            Console.Write("Exit now? ");
            do
            {
                Console.Write("Enter your choice (YES / NO): ");
                choice = Console.ReadLine().ToLower();

                if (choice == "yes")
                {
                    check = 1;
                    Console.WriteLine("Goodbye!");
                    Environment.Exit(0);
                }
                else if (choice == "no")
                {
                    check = 1;
                    Console.WriteLine("Attempt to go back to Main Menu. Press any keys to proceed...");
                    Console.ReadKey();
                }
                else
                {
                    Console.WriteLine("Error: Invalid choice. Please try again.");
                }
            } while (check == 0);
        }

        public void ExportToExcel()
        {
            int count = students.Count();

            if (count > 0)
            {
                ExcelPackage excelPackage = new ExcelPackage();

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Create headings
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "First Name";
                worksheet.Cells[1, 3].Value = "Last Name";
                worksheet.Cells[1, 4].Value = "Email";
                worksheet.Cells[1, 5].Value = "Address";
                worksheet.Cells[1, 6].Value = "Phone Number";

                int row = 2;
                foreach (var student in students)
                {
                    worksheet.Cells[row, 1].Value = student.ID;
                    worksheet.Cells[row, 2].Value = student.FirstName;
                    worksheet.Cells[row, 3].Value = student.LastName;
                    worksheet.Cells[row, 4].Value = student.Email;
                    worksheet.Cells[row, 5].Value = student.Address;
                    worksheet.Cells[row, 6].Value = student.PhoneNumber;
                    row++;
                }

                string filePath = @"C:\Users\nggng\Documents\FPT\Code\NPL.ASM12\NPL.ASM12\bin\Debug\net6.0\students.xlsx";

                // Delete file if it is already existed
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }

                FileStream fileStream = File.Create(filePath);
                fileStream.Close();

                File.WriteAllBytes(filePath, excelPackage.GetAsByteArray());

                excelPackage.Dispose();

                Console.WriteLine("Data exported to Excel file successfully. Press any key to proceed...");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Error: No data in list found. Press any keys to proceed...");
                Console.ReadKey();
            }
        }

        public void ImportFromExcel()
        {
            string filePath = @"C:\Users\nggng\Documents\FPT\Code\NPL.ASM12\NPL.ASM12\bin\Debug\net6.0\students.xlsx";

            if (File.Exists(filePath))
            {
                ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath));

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Sheet1"];

                int rowCount = worksheet.Dimension.Rows;

                // Skip the first row as it contains the headings
                for (int row = 2; row <= rowCount; row++)
                {
                    Student student = new Student();

                    student.ID = worksheet.Cells[row, 1].Value?.ToString();
                    student.FirstName = worksheet.Cells[row, 2].Value?.ToString();
                    student.LastName = worksheet.Cells[row, 3].Value?.ToString();
                    student.Email = worksheet.Cells[row, 4].Value?.ToString();
                    student.Address = worksheet.Cells[row, 5].Value?.ToString();
                    student.PhoneNumber = worksheet.Cells[row, 6].Value?.ToString();

                    students.Add(student);
                }

                excelPackage.Dispose();

                Console.WriteLine("Data imported from Excel file successfully. Press any key to proceed...");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Error: Excel file not found. Press any keys to proceed...");
                Console.ReadKey();
            }
        }
    }
}
