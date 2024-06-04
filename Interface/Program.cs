public partial class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Please select an operation to run, enter the number from he list:");
        Console.WriteLine("1. Excelhandler");

        ConsoleKeyInfo Selection = Console.ReadKey();

        Console.WriteLine("Please confirm that you want to execute the selected operation (" + Selection.Key + ") y/n");

        ConsoleKeyInfo ConfirmationKey = Console.ReadKey();
        if (ConfirmationKey.Key == ConsoleKey.Y)
        {
            Console.WriteLine("Operation confirmed... Starting Process");
        }
        else
        {
            Console.Clear();
            Console.WriteLine("Operation cancelled... Lets try again");
            Main(args);
        }

        try
        {
            switch (Selection.Key)
            {
                case ConsoleKey.D1:
                case ConsoleKey.NumPad1:
                    Console.WriteLine("Please enter the fully qualified file path of the excel sheet you want to use as a test subject");
                    var FilePath = Console.ReadLine();
                    if (FilePath != null)
                    {
                        ExcelHandler.Read.ReadToDataSet(FilePath);
                    }

                    break;
                default:
                    Console.Clear();
                    Console.WriteLine("The selection made does not exist, please choose an operation number from the list provided");
                    Main(args);
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Main(args);
        }
    }
}

