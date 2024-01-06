namespace NPL.ASM12
{
    public static class InputHelper
    {
        public static bool InputInteger(int min, int max, out int result)
        {
            while (true)
            {
                bool isValid = int.TryParse(Console.ReadLine(), out result);
                if (isValid && result > min && result <= max)
                {
                    return isValid;
                }
                else
                {
                    return false;
                }
            }
        }

        public static int CreateMenu(string menuTitle, List<string> menuItems)
        {
            int choice;
            while (true)
            {
                Console.Clear();
                Console.WriteLine($"========= {menuTitle} =========");
                menuItems.ForEach(item => Console.WriteLine($"{menuItems.IndexOf(item) + 1}. {item}"));

                Console.Write("Enter a choice: ");
                bool rightChoice = InputHelper.InputInteger(0, menuItems.Count, out choice);

                if (rightChoice)
                {
                    break;
                }
                else
                {
                    Console.WriteLine("Invalid choice. Press any key to try again...");
                    Console.ReadKey();
                }
            }

            return choice;
        }
    }
}
