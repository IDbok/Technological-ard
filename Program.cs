using OfficeOpenXml;
using System.Text.Json;

namespace Technological_card
{
    internal class Program
    {
        // Путь к файлу с ТК
        static string filepath = @"C:\Users\bokar\Documents\ТК.xlsx";

        //Список ключевых слов (заголовков)
        static string[] keyWords =
        {
                "1. Требования к составу бригады и квалификации",
                "2. Требования к материалам и комплектующим",
                "3. Требования к механизмам",
                "4. Требования к средствам защиты",
                "5. Требования к инструментам и приспособлениям",
                "6. Выполнение работ"

            };
        static string[] stuctNames = { "Staff", "Components", "Machines", "Protection", "Tools" };
        static int numTitle = 5 - 1;

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            // Ввод лицензии для работы с Excel
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            var dicFirstRows = new Dictionary<string, int>();

            //Заполняем словарь ключевыми словами
            foreach (string st in keyWords)
            {
                dicFirstRows.Add(st, 1);
            }
            
            // Создаю объект для работы с Excel
            using (var package = new ExcelPackage(new FileInfo(filepath)))
            {
                //Получение названия всех листов с ТК в книге
                List<String> sheetsTK = new();
                foreach (var ws in package.Workbook.Worksheets)
                {
                    if (ws.Name.ToString().Contains("ТК_")) sheetsTK.Add(ws.Name);
                }

                // TODO: Цыкл для всех листов с ТК

                // Определение листа в переменную
                string sheetName = sheetsTK[3];
                var worksheet = package.Workbook.Worksheets[sheetName];

                // Поиск номеров строк с ключевыми словами. Запись их в словарь
                int numRowsFound = 0;
                for (int i = 1; i < 1000; i++)
                {
                    string? valueCell = worksheet.Cells[i, 1].Value.ToString();
                    string keyWord = keyWords[numRowsFound];
                    if (valueCell == keyWord)
                    {
                        dicFirstRows[keyWord] = i;
                        numRowsFound++;
                    }
                    if (numRowsFound == dicFirstRows.Count) { break; }
                }

                // Создание списков объектов моделей данных
                //List<Staff> Staffs = new();
                //List<Tool> Tools = new();
                //List<Machine> Machines = new();
                //List<Protection> Protections = new();
                var Components = new List<Struct>();

                // Запись данных из Excel в список в соответствии с моделью данных
                int startRow = dicFirstRows[keyWords[numTitle]] + 2;
                int endRow = dicFirstRows[keyWords[numTitle + 1]];
                for (int i = startRow; i < endRow; i++)
                {
                    string num = worksheet.Cells[i, 1].Value.ToString().Trim();
                    string name = worksheet.Cells[i, 2].Value.ToString().Trim();
                    string type = worksheet.Cells[i, 3].Value.ToString().Trim();
                    string unit = worksheet.Cells[i, 4].Value.ToString().Trim();
                    string amount = worksheet.Cells[i, 5].Value.ToString().Trim();

                    Components.Add(new Component
                    {
                        Num = int.Parse(num),
                        Name = name,
                        Type = type,
                        Unit = unit,
                        Amount = (int)uint.Parse(amount)
                    });
                    
                }

                SaveToJSON(Components, sheetName);
                
                Console.WriteLine($"Парсиг карты на листе {sheetName} законцен!");
            }
        }

        public static void SaveToJSON(List<Struct> ListOfStruct, string sheetName) 
        {
            foreach (var item in ListOfStruct)
            {
                // Создание объекта для записи без кодировки в Unicode
                var options = new JsonSerializerOptions
                {
                    // set the encoder to UnicodeEncoding
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string json = JsonSerializer.Serialize(item, options);
                string pathJson = $@"C:\Users\bokar\source\repos\Technological card\Cards\{sheetName}\{stuctNames[numTitle]}\";
                if (!Directory.Exists(pathJson)) Directory.CreateDirectory(pathJson);
                File.WriteAllText(pathJson + item.Num + ".json", json);
            }
        }
        public static void SaveToJSON(List<Staff> ListOfStruct, string sheetName)
        {
            foreach (var item in ListOfStruct)
            {
                // Создание объекта для записи без кодировки в Unicode
                var options = new JsonSerializerOptions
                {
                    // set the encoder to UnicodeEncoding
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string json = JsonSerializer.Serialize(item, options);
                string pathJson = $@"C:\Users\bokar\source\repos\Technological card\Cards\{sheetName}\Components\";
                if (!Directory.Exists(pathJson)) Directory.CreateDirectory(pathJson);
                File.WriteAllText(pathJson + item.Num + ".json", json);
            }
        }
        public static void SaveToJSON(List<Machine> ListOfStruct, string sheetName)
        {
            foreach (var item in ListOfStruct)
            {
                // Создание объекта для записи без кодировки в Unicode
                var options = new JsonSerializerOptions
                {
                    // set the encoder to UnicodeEncoding
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string json = JsonSerializer.Serialize(item, options);
                string pathJson = $@"C:\Users\bokar\source\repos\Technological card\Cards\{sheetName}\Components\";
                if (!Directory.Exists(pathJson)) Directory.CreateDirectory(pathJson);
                File.WriteAllText(pathJson + item.Num + ".json", json);
            }
        }
    }
}