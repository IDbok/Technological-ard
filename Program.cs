using CsvHelper;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Technological_card
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            //string inputfilepath = "C:\Users\bokar\source\repos\Technological card\Components.csv";

            //List<ComponentModel> components = new();

            //using var reader = new StringReader(inputfilepath);
            //using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

            //var records = csv.GetRecords<ComponentModel>().ToList();

            //foreach (var record in records)
            //{
            //    Console.WriteLine(record.num); 
            //}

            //string inputfilepath = "C:\\Users\\bokar\\source\\repos\\Technological card\\Components.csv";

            //using var reader = new StreamReader(inputfilepath);
            //using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

            //var records = csv.GetRecords<ComponentModel>().ToList();

            //foreach (var record in records)
            //{
            //    Console.WriteLine(record.num);
            //}


            Component newcomp = NewComp();

            Console.WriteLine("Действительно создан объект" + newcomp.ToString());

            //string filePath = "C:\\Users\\bokar\\OneDrive\\Работа\\Таврида\\Технологические карты\\Пример\\Копия Технологические карты.xlsx";


            //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            //{
            //    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            //    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            //    Worksheet worksheet = worksheetPart.Worksheet;
            //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();


            //    foreach (Row row in sheetData.Elements<Row>())
            //    {
            //        foreach (Cell cell in row.Elements<Cell>())
            //        {
            //            string cellValue = cell.CellValue.Text;
            //            // Обработка данных
            //            Console.WriteLine(cellValue);
            //            Console.ReadLine();
            //        }
            //    }
            //}


        }



        public static Component NewComp()
        {
            Component comp = new Component();
            //string outputString;
            int outputInt;
            bool[] outputCheck = new bool[2];
            bool finalCheck = false;
            while (finalCheck == false)
            {
                Console.WriteLine("Введите позицию инструмента или приспособления: ");
                outputCheck[0] = int.TryParse(Console.ReadLine(), out outputInt);
                comp.Num = outputInt;

                Console.WriteLine("Наименование: ");
                comp.Name = Console.ReadLine();

                Console.WriteLine("Тип (исполнение): ");
                comp.Type = Console.ReadLine();

                Console.WriteLine("Ед. измерения: ");
                comp.Unit = Console.ReadLine();

                Console.WriteLine("Количество: ");
                outputCheck[1] = int.TryParse(Console.ReadLine(), out outputInt);
                outputCheck[1] = outputInt != 0;
                comp.Amount = outputInt;


                finalCheck = outputCheck[1] && outputCheck[1];
            }
            Console.WriteLine("Создан новый объект инструментов:\n" + comp.ToString());
            return comp;
        }
    }
}