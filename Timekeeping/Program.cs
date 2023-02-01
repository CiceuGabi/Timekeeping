using ClosedXML.Excel;

namespace Timekeeping;

public class Program
{
    public static void Main(string[] args)
    {
        EditExcel();
    }

    private static void EditExcel()
    {
        var workbook = new XLWorkbook(Environment.CurrentDirectory + @"\Model.xlsx");

        var worksheet = workbook.Worksheet(1);
        var cell = worksheet.Cell("F7").Value;

        /*
         Enter your path below where you want to save the excel file.
         The path should be something like this:  @"C:\Users\ciceu"   
        */
        var userPath = @"C:\Users\ciceu";

        if (userPath == string.Empty)
        {
            Console.WriteLine("You need to enter an output path");
            Environment.Exit(0);
        }

        Console.WriteLine();

        var userInput = GetMonthName();
        var outputPath = $@"{userPath}\Luna {userInput} 2023.xlsx";

        if (cell.ToString().Contains("{MONTH}"))
        {
            worksheet.Cell("F7").Value = $"LUNA {userInput} 2023";
            worksheet.Name = userInput + " 2023";
            FindDaysInMonth(userInput, worksheet);
        }

        Console.WriteLine("Month {0} created successfully", userInput);
        workbook.SaveAs(outputPath);
    }

    private static void FindDaysInMonth(string userInput, IXLWorksheet worksheet)
    {
        var month = GetMonthNumber(userInput);
        DateTime date = new DateTime(DateTime.Now.Year, month, 1);
        int daysInMonth = DateTime.DaysInMonth(date.Year, date.Month);
        int column = 4;

        for (int i = 0; i < daysInMonth; i++)
        {
            worksheet.Cell(10, column).Value = date.Day;
            var dayOfWeek = date.DayOfWeek.ToString();
            CheckForWeekends(dayOfWeek, worksheet, column);
            CheckForNationalDays(date.Day, userInput, worksheet, column);
            date = date.AddDays(1);
            column++;
        }

        DeleteUnusedColumns(daysInMonth, worksheet);
    }

    private static void DeleteUnusedColumns(int daysInMonth, IXLWorksheet worksheet)
    {
        switch (daysInMonth)
        {
            case 28:
                worksheet.Column(34).Delete();
                worksheet.Column(33).Delete();
                worksheet.Column(32).Delete();
                worksheet.Range(9, 19, 9, 31).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                break;
            case 29:
                worksheet.Column(34).Delete();
                worksheet.Column(33).Delete();
                worksheet.Range(9, 19, 9, 32).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                break;
            case 30:
                worksheet.Column(34).Delete();
                worksheet.Range(9, 19, 9, 33).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                break;
        }
    }

    private static void CheckForWeekends(string dayOfWeek, IXLWorksheet worksheet, int limit)
    {
        switch (dayOfWeek)
        {
            case "Sunday":
            case "Saturday":
                CellEdit(worksheet, limit);
                break;
        }
    }

    private static void CheckForNationalDays(int day, string userInput, IXLWorksheet worksheet, int limit)
    {
        List<int> ianuarie = new List<int> { 1, 2, 24 };
        List<int> aprilie = new List<int> { 14, 16, 17 };
        List<int> mai = new List<int> { 1 };
        List<int> iunie = new List<int> { 1, 4, 5 };
        List<int> august = new List<int> { 15 };
        List<int> noiembrie = new List<int> { 30 };
        List<int> decembrie = new List<int> { 1, 25, 26 };

        Dictionary<string, List<int>> nationalDays = new Dictionary<string, List<int>>
        {
            { "IANUARIE", ianuarie }, { "APRILIE", aprilie }, { "MAI", mai }, { "IUNIE", iunie },
            { "AUGUST", august }, { "NOIEMBRIE", noiembrie }, { "DECEMBRIE", decembrie }
        };

        if (nationalDays.ContainsKey(userInput) && nationalDays[userInput].Contains(day))
        {
            CellEdit(worksheet, limit);
        }
    }

    private static void CellEdit(IXLWorksheet worksheet, int limit)
    {
        worksheet.Range(10, limit, 16, limit).Style.Fill.BackgroundColor = XLColor.Yellow;
        worksheet.Column(limit).Width = 2.2;
        worksheet.Cell(11, limit).Style.Font.FontSize = 11;
        worksheet.Cell(11, limit).Value = "L";
    }

    private static void DisplayMenu()
    {
        Dictionary<string, string> menuContent = new Dictionary<string, string>
        {
            { "1", "IANUARIE" }, { "2", "FEBRUARIE" }, { "3", "MARTIE" }, { "4", "APRILIE" }, { "5", "MAI" }, { "6", "IUNIE" },
            { "7", "IULIE" }, { "8", "AUGUST" }, { "9", "SEPTEMBRIE" }, { "10", "OCTOMBRIE" }, { "11", "NOIEMBRIE" }, { "12", "DECEMBRIE" }
        };

        Console.WriteLine("Choose one of the following: ");
        Console.WriteLine("---------------");
        foreach (var menu in menuContent)
        {
            Console.WriteLine(menu.Key + " - " + menu.Value);
        }

        Console.WriteLine("---------------");
    }

    private static int GetMonthNumber(string input)
    {
        Dictionary<string, int> months = new Dictionary<string, int>
        {
            { "IANUARIE", 1 }, { "FEBRUARIE", 2 }, { "MARTIE", 3 }, { "APRILIE", 4 }, { "MAI", 5 }, { "IUNIE", 6 },
            { "IULIE", 7 }, { "AUGUST", 8 }, { "SEPTEMBRIE", 9 }, { "OCTOMBRIE", 10 }, { "NOIEMBRIE", 11 }, { "DECEMBRIE", 12 }
        };
        var inputNum = 0;
        foreach (var month in months.Where(m => input.Equals(m.Key)))
        {
            inputNum = month.Value;
        }

        return inputNum;
    }

    private static string GetMonthName()
    {
        DisplayMenu();
        Console.Write("Please enter your number: ");
        var month = GetInputNumber() switch
        {
            1 => "IANUARIE",
            2 => "FEBRUARIE",
            3 => "MARTIE",
            4 => "APRILIE",
            5 => "MAI",
            6 => "IUNIE",
            7 => "IULIE",
            8 => "AUGUST",
            9 => "SEPTEMBRIE",
            10 => "OCTOMBRIE",
            11 => "NOIEMBRIE",
            12 => "DECEMBRIE",
            _ => ""
        };

        return month;
    }

    private static int GetInputNumber()
    {
        while (true)
        {
            var input = Console.ReadLine();

            if (int.TryParse(input, out var num))
            {
                if (num is >= 1 and <= 12)
                {
                    return num;
                }

                Console.Write("Number must be between 1-12: ");
            }
            else
            {
                Console.Write("This is not a number. Enter a valid number: ");
            }
        }
    }
}