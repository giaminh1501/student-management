using System.Text.RegularExpressions;

namespace NPL.ASM12
{
    public class Student
    {
        public string ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }

        public Student(string id, string firstName, string lastName, string email, string address, string phoneNumber)
        {
            ID = id;
            FirstName = firstName;
            LastName = lastName;
            Email = email;
            Address = address;
            PhoneNumber = phoneNumber;
        }

        public Student()
        {

        }

        public void CheckID()
        {
            if (!Regex.IsMatch(ID, @"^[a-zA-Z0-9]+$"))
            {
                Console.WriteLine("Error: Invalid ID. Must not conatain special characters. Must be unique. Only free characters and numbers are accepted.");
            }
        }

        public void CheckFirstName()
        {
            if (!Regex.IsMatch(FirstName, @"^[a-zA-Z\s]+$"))
            {
                Console.WriteLine("Error: Invalid first name format. Must not contain numbers or special characters.");
            }
        }

        public void CheckLastName()
        {
            if (!Regex.IsMatch(LastName, @"^[a-zA-Z\s]+$"))
            {
                Console.WriteLine("Error: Invalid last name format. Must not contain numbers or special characters.");
            }
        }

        public void CheckEmail()
        {
            if (!Regex.IsMatch(Email, @"^[a-zA-Z0-9./]+@(outlook\.com|gmail\.com)$"))
            {
                Console.WriteLine("Error: Invalid email format. User name part must contain only free characters, numbers or special characters like '.' or '/'. Domain name part must be 'outlook.com' or 'gmail.com'.");
            }
        }

        public void CheckPhoneNumber()
        {
            if (!Regex.IsMatch(PhoneNumber, @"^(?:\+84|0)[0-9]{9}$"))
            {
                Console.WriteLine("Error: Invalid phone number format. Must be started with '+84' or '0', followed by 9 degits number.");
            }
        }
    }
}

