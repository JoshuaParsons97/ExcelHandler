using Interface;

public partial class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Please select an operation to run, enter the number from he list:");
        Console.WriteLine("1. Excelhandler Read");
        Console.WriteLine("2. Excelhandler Write (uses a test data set)");

        ConsoleKeyInfo Selection = Console.ReadKey();
        Console.WriteLine("");

        Console.WriteLine("Please confirm that you want to execute the selected operation (" + Selection.Key + ")");
        Console.WriteLine("y/n?");

        ConsoleKeyInfo ConfirmationKey = Console.ReadKey();
        Console.Clear();
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

        Console.WriteLine("");

        try
        {
            switch (Selection.Key)
            {
                case ConsoleKey.D1:
                case ConsoleKey.NumPad1:
                    Console.WriteLine("Please enter the fully qualified file path of the excel sheet you want to use as a test subject");
                    var ReadFilePath = Console.ReadLine();
                    if (ReadFilePath != null)
                    {
                        ExcelHandler.Read.ReadToDataSet(ReadFilePath);
                    }

                    break;
                case ConsoleKey.D2:
                case ConsoleKey.NumPad2:
                    Console.WriteLine("This process will execute using a test data set as a test subject, please enter the filepath where you want to save the excel document");
                    var SaveFilePath = Console.ReadLine();
                    if(SaveFilePath != null)
                    {
                        ExcelHandler.Write.DatasetToExcel(TestData.TestDataSet, SaveFilePath);
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

