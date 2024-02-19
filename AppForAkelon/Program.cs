using System;
using ClosedXML.Excel;
using System.IO;
using System.Linq;


namespace AppForAkelon
{
    public class Product
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }

        // Конструктор
        public Product(string code, string name, double price)
        {
            Code = code;
            Name = name;
            Price = price;
        }

        // Метод для отображения информации о товаре
        public void DisplayInfo()
        {
            Console.WriteLine($"Код товара: {Code}");
            Console.WriteLine($"Цена за единицу: {Price:F2} руб");
        }

        // Метод для проверки соответствия товара по наименованию
        public bool MatchesName(string productName)
        {
            return Name == productName;
        }
    }

    public class Client
    {
        public string OrganizationName { get; set; }
        public string ContactPerson { get; set; }

        public Client(string organizationName, string contactPerson)
        {
            OrganizationName = organizationName;
            ContactPerson = contactPerson;
        }

        public bool MatchesOrganizationName(string orgName)
        {
            return OrganizationName == orgName;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите путь до Excel документа\nПример: C:\\Users\\user\\name_file.xlsx");
            string filePath = Console.ReadLine();
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new Exception("Неверный путь или название файла, убедитесь что введен правильный путь или название файла");
                }

                using (var workbook = new XLWorkbook(filePath))
                {

                    while (true)
                    {
                        Console.WriteLine("\tЧто нужно сдлеать:");
                        Console.WriteLine("1. Получить информацию по товару");
                        Console.WriteLine("2. Изменить контактное лицо");
                        Console.WriteLine("3. Получить \"золотого клиента\"");
                        Console.WriteLine("0. Выход");

                        int choice = int.Parse(Console.ReadLine());

                        switch (choice)
                        {
                            case 1:
                                Console.WriteLine("Введите название товара: наприер \"Молоко\"");
                                string prodName = Console.ReadLine();
                                SearchProd(workbook, prodName);
                                break;

                            case 2:
                                Console.WriteLine("Введите название организации: например ООО Надежда");
                                string orgName = Console.ReadLine();
                                Console.WriteLine("Введите ФИО нового контактного лица:");
                                string newClientName = Console.ReadLine();
                                NewClient(workbook, orgName, newClientName);
                                break;

                            case 3:
                                Console.WriteLine("Введите год:");
                                int year = int.Parse(Console.ReadLine());
                                Console.WriteLine("Введите месяц:");
                                int month = int.Parse(Console.ReadLine());
                                Console.WriteLine("\n");
                                SearchGoldClient(workbook, year, month);
                                Console.WriteLine("\n");
                                break;

                            case 0:
                                return;

                            default:
                                Console.WriteLine("Повторите попытку");
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}");
            }
            Console.ReadKey();

            //метод поиска инфы по товару
            void SearchProd(XLWorkbook workbook, string prodName)
            {
                try
                {
                    var prodSheet = workbook.Worksheet(1); // Первый лист с товарами
                    var clientSheet = workbook.Worksheet(2); // Второй лист с клиентами
                    var ordSheet = workbook.Worksheet(3); // Третий лист с заказами

                    var prodRows = prodSheet.RowsUsed().Skip(1); // Получаем строки, пропуская первую

                    bool prodFound = false;

                    foreach (var row in prodRows)
                    {
                        // Создаем экземпляр класса Product для текущей строки
                        var product = new Product(row.Cell("A").GetString(), row.Cell("B").GetString(), row.Cell("D").GetDouble());

                        if (product.Name == prodName)
                        {
                            Console.WriteLine($"\tТовар '{prodName}'");
                            Console.WriteLine($"Код товара: {product.Code}");
                            Console.WriteLine($"Цена за единицу: {product.Price:F2} руб");

                            var orderRows = ordSheet.RowsUsed().Skip(1);

                            foreach (var orderRow in orderRows)
                            {
                                var ordProdCode = orderRow.Cell("B").GetString();
                                var ordDate = orderRow.Cell("F").GetString();

                                if (ordProdCode == product.Code)
                                {
                                    var clientCode = orderRow.Cell("C").GetString();

                                    var clientRows = clientSheet.RowsUsed().Skip(1);

                                    foreach (var clientRow in clientRows)
                                    {
                                        if (clientRow.Cell("A").GetString() == clientCode)
                                        {
                                            Console.WriteLine($"Клиент: {clientRow.Cell("B").GetString()}");
                                            Console.WriteLine($"Адрес: {clientRow.Cell("C").GetString()}");
                                            Console.WriteLine($"Количество товара: {orderRow.Cell("C").GetDouble()}");
                                            Console.WriteLine($"Дата размещения: {ordDate}");
                                            Console.WriteLine("\n");
                                            prodFound = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            break;
                        }
                    }

                    if (!prodFound)
                    {
                        Console.WriteLine($"Товар {prodName} не найден.");
                        Console.WriteLine("\n");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                }
            }

            //метод изменения клиента
            void NewClient(XLWorkbook workbook, string orgName, string newClientName)
            {
                try
                {
                    var clientSheet = workbook.Worksheet(2);

                    var clientRows = clientSheet.RowsUsed().Skip(1);

                    bool clientFound = false;
                    foreach (var clientRow in clientRows)
                    {
                        var client = new Client(clientRow.Cell("B").GetString(), clientRow.Cell("D").GetString());

                        if (client.OrganizationName == orgName)
                        {
                            // Изменяем свойство
                            client.ContactPerson = newClientName;

                            clientRow.Cell("D").Value = newClientName;

                            clientFound = true;
                            if (clientFound)
                                Console.WriteLine("ФИО установлено :)");
                            else
                                Console.WriteLine("Такой организации нет :(");
                            break;
                        }
                    }
                    workbook.Save();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                }
            }

            //метод поиска золотого клиента
            void SearchGoldClient(XLWorkbook workbook, int year, int month)
            {
                try
                {
                    var clientSheet = workbook.Worksheet(2);
                    var ordSheet = workbook.Worksheet(3);

                    var ordRows = ordSheet.RowsUsed().Skip(1);
                    var searchOrdDate = ordRows.Where(row =>
                    {
                        var ordDate = row.Cell("F").GetDateTime();
                        return ordDate.Year == year && ordDate.Month == month;
                    });

                    var groupCodeClient = searchOrdDate.GroupBy(row => row.Cell("C").GetString()); //присвоение группы

                    string goldClient = "";
                    var maxOrders = 0;

                    foreach (var group in groupCodeClient)
                    {
                        var ordCount = group.Count();

                        if (ordCount > maxOrders)
                        {
                            maxOrders = ordCount;
                            goldClient = group.Key;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(goldClient)) //учет пробелов пусой строки
                    {
                        var topClient = clientSheet.RowsUsed().FirstOrDefault(row => row.Cell("A").GetString() == goldClient);
                        if (topClient != null)
                        {
                            Console.WriteLine($"Золотой клиент за {month}/{year}: {topClient.Cell("B").GetString()}");
                            Console.WriteLine($"Количество заказов: {maxOrders}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Золотой клиент не найден.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message}");
                }
            }
        }
    }
}








