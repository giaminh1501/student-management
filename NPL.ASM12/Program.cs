namespace NPL.ASM12;

public class Program
{
    static public void Main(string[] args)
    {
        StudentManagement studentManagement = new StudentManagement();

        Student student = new Student();

        MainMenu();

        void MainMenu()
        {
            Console.Clear();
            string menuTitle = "Student Management System";
            var menuItems = new List<string>()
            {
                "Add student",
                "Remove a student",
                "Display all students",
                "Export data to Excel file",
                "Read data from Excel file",
                "Exit"
            };

            int choice = InputHelper.CreateMenu(menuTitle, menuItems);

            switch (choice)
            {
                case 1:
                    Console.Clear();
                    studentManagement.Add();

                    MainMenu();

                    break;

                case 2:
                    Console.Clear();
                    studentManagement.Remove();

                    MainMenu();

                    break;

                case 3:
                    Console.Clear();
                    studentManagement.Display();

                    MainMenu();

                    break;

                case 4:
                    Console.Clear();
                    studentManagement.ExportToExcel();

                    MainMenu();

                    break;

                case 5:
                    Console.Clear();
                    studentManagement.ImportFromExcel();

                    MainMenu();

                    break;

                case 6:
                    Console.Clear();                   
                    studentManagement.CheckExit();

                    MainMenu();
                
                    break;
            }
        }
    }
}
