using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Akelon.Models;
using ClosedXML;
using ClosedXML.Excel;

namespace Akelon
{
    public class Program
    {
        static HashSet<Client> clients = new HashSet<Client>();
        static HashSet<Request> requests = new HashSet<Request>();
        static HashSet<Ware> wares = new HashSet<Ware>();

        /// <summary>
        /// Запрос на ввод пути до файла с данными
        /// </summary>
        static string GetFileName()
        {
            // считываем путь до файла с данными
            Console.WriteLine("Введите путь до файла с данными или оставьте пустым, если данные находятся в одной папке с данной программой: ");
            var filePath = Console.ReadLine()!.Trim()!;
            if (filePath == "")
                filePath = Directory.GetCurrentDirectory();

            // проверка на существование директории
            while (!Directory.Exists(filePath))
            {
                Console.WriteLine("Введена некорректная директория, повторите ввод." +
                    "\n\tОставьте путь пустым, если данные находятся в одной папке с программой");
                filePath = Console.ReadLine()!.Trim()!;
            }

            // считываем файл с расширением
            Console.WriteLine("Введите имя файла с расширением: ");
            var fileName = Console.ReadLine()!.Trim()!;
            while (!File.Exists(Path.Combine(filePath, fileName)))
            {
                Console.WriteLine("Введено некорректное имя файла или его расширение. Попробуйте еще раз: ");
                fileName = Console.ReadLine()!.Trim()!;
            }

            return Path.Combine(filePath, fileName);
        }

        /// <summary>
        /// Считываем данные из файла и заполняем сеты данными
        /// </summary>
        /// <param name="fname">Имя файла с данными</param>
        static void FillSetsFromFile(string fname)
        {
            using (var workbook = new XLWorkbook(fname))
            {
                wares = workbook.Worksheet("Товары")
                    .Rows()
                    .Where(row => row.FirstCell().Value.IsNumber)
                    .Select(row => new Ware(row))
                    .ToHashSet();

                clients = workbook.Worksheet("Клиенты")
                    .Rows()
                    .Where(row => row.FirstCell().Value.IsNumber)
                    .Select(row => new Client(row))
                    .ToHashSet();

                requests = workbook.Worksheet("Заявки")
                    .Rows()
                    .Where(row => row.FirstCell().Value.IsNumber)
                    .Select(row => new Request(row))
                    .ToHashSet();
            }
        }

        /// <summary>
        /// По наименованию товара выводить информацию о клиентах, заказавших этот товар, с указанием информации по 
        /// количеству товара, цене и дате заказа.
        /// </summary>
        static void FindClientsByWareName()
        {
            Console.WriteLine("\nВведите название товара для поиска: ");
            var wareName = Console.ReadLine()!.Trim()!;
            // обработка - товар не должен быть пустым
            while(wareName == "")
            {
                Console.WriteLine("Название товара не может быть пустым. Введите еще раз: ");
                wareName = Console.ReadLine()!.Trim()!;
            }

            Console.WriteLine($"\nТовар \"{wareName}\" заказали клиенты:\n");
            var wareCode = wares.First(ware => ware.Name == wareName).Code;

            requests
                .Where(request => request.WareCode == wareCode)
                .Select((request, i) => 
                {
                    var client = clients.First(client => client.Code == request.ClientCode);
                    return $"Клиент #{i}:\n\t{client}\nЕго заявка:\n\t{request}\n";
                })
                .ToList()
                .ForEach(result => Console.WriteLine(result));
        }

        /// <summary>
        /// Запрос на определение золотого клиента
        /// </summary>
        static void FindGoldClient()
        {
            Console.WriteLine("\nПоиск золотого клиента...");
            var goldRequire = requests
                .GroupBy(request => request.ClientCode)
                .OrderByDescending(request => request.Sum(x => x.Count))
                .First();

            var goldClient = clients.First(client => client.Code == goldRequire.First().ClientCode);
            Console.WriteLine($"Золотой клиент:\n\t{goldClient}\nУ него наибольшее число покупок: {goldRequire.Sum(x => x.Count)}");
        }

        /// <summary>
        /// Запрос на определение клиента с наибольшим количеством заказов, за указанный год, месяц.
        /// </summary>
        static void FindMaxRequiresClient()
        {
            Console.WriteLine("\nВведите год, за который необходимо определить клиента с наибольшим количеством заказов: ");
            int year = 0;
            while(!int.TryParse(Console.ReadLine()!.Trim()!, out year))
            {
                Console.WriteLine("Введен некорректный год, попробуйте еще раз: ");
            }

            Console.WriteLine("Введите номер месяца(число от 1 до 12 включительно): ");
            int month = -1;
            while (!int.TryParse(Console.ReadLine()!.Trim()!, out month) || (month <= 0 || month >= 13))
            {
                Console.WriteLine("Введен некорректный месяц, попробуйте еще раз: ");
            }

            Console.WriteLine("Начинаю поиск...");
            var maxRequires = requests
                .Where(request => request.Date.Month == month && request.Date.Year == year)
                .GroupBy(request => request.ClientCode)
                .OrderByDescending(request => request.Sum(x => x.Count))
                .FirstOrDefault();

            if (maxRequires != null)
            {
                var client = clients.First(client => client.Code == maxRequires.First().ClientCode);
                Console.WriteLine($"Клиент:\n\t{client}\nНабрал наибольшее число заказов({maxRequires.Sum(x => x.Count)}) " +
                    $"за {month.MonthFromInt()} {year}");
            }
            else
                Console.WriteLine($"За {month.MonthFromInt()} {year} не найдено ни одного заказов");
        }

        /// <summary>
        /// Запрос на изменение контактного лица клиента с указанием параметров: Название организации, ФИО нового контактного лица. 
        /// В результате информация должна быть занесена в этот же документ, в качестве ответа пользователю необходимо выдавать
        /// информацию о результате изменений.
        /// </summary>
        static void ChangeContactPerson(string fname)
        {
            // считываем организацию
            Console.WriteLine("\nВведите название организации: ");
            var companyName = Console.ReadLine()!.Trim();
            // обработка введенного результата - проверяем, что такая организация существует
            while (clients.Count(client => client.Name == companyName) == 0)
            {
                Console.WriteLine("Введено некорректное название организации. Попробуйте ввести еще раз: ");
                companyName = Console.ReadLine()!.Trim()!;
            }

            // считываем новое контактное лицо
            Console.WriteLine("Введите ФИО нового контактного лица: ");
            var newContact = Console.ReadLine()!.Trim()!;
            // обработка результата - проверяем, что контакт не пустой
            while(newContact == "")
            {
                Console.WriteLine("ФИО нового контактно лица введено некорректно. Попробуйте ввести еще раз: ");
                newContact = Console.ReadLine()!.Trim()!;
            }

            // находим нужного клиента в программе и меняем данные
            var client = clients.First(client => client.Name == companyName)!;
            client.ContactPerson = newContact;

            // сохраняем данные в файле
            using (var workbook = new XLWorkbook(fname))
            {
                workbook
                    .Worksheet("Клиенты")
                    .Rows()
                    .First(row =>
                    {
                        var firstCell = row.FirstCell().Value;
                        return firstCell.IsNumber && (int)firstCell.GetNumber() == client.Code;
                    })
                    .Cell("D")
                    .SetValue(newContact);
                workbook.Save();
            }

            // выводим пользователю, что данные обновлены и сохранены
            Console.WriteLine("Изменения успешно сохранены в программе и занесены в файл!");
        }

        static void Main(string[] args)
        {
            /*
            Разработать консольное приложение на языке С#, которое будет выполнять следующие команды:
                1. Запрос на ввод пути до файла с данными (в качестве документа с данными использовать Приложение 2).
                2. По наименованию товара выводить информацию о клиентах, заказавших этот товар, с указанием информации по 
                    количеству товара, цене и дате заказа.
                3. Запрос на изменение контактного лица клиента с указанием параметров: Название организации, ФИО нового контактного лица. 
                    В результате информация должна быть занесена в этот же документ, в качестве ответа пользователю необходимо выдавать 
                    информацию о результате изменений.
                4. Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год, месяц.
            Для работы с документом рекомендуем использовать свободно распространяемую библиотеку OpenXML или ClosedXML.
             */

            Console.WriteLine("Начало работы...");
            var fname = GetFileName();
            Console.WriteLine("Загрузка данных из файла...");
            FillSetsFromFile(fname);
            Console.WriteLine("Данные успешно загружены.\n");

            var exit = false;
            while (!exit)
            {
                Console.WriteLine("\n============================================================\n");
                Console.WriteLine("Что требуется сделать с файлом? Нажмите" + 
                    "\n\t1 - если нужно вывести информацию о клиентах и об их заказе по названию товара, который они заказали;" + 
                    "\n\t2 - если нужно изменить контактное лицо клиента по названию организации;" + 
                    "\n\t3 - если нужно узнать золотого клиента;" + "" +
                    "\n\t4 - если нужно узнать клиента с наибольшим количеством заказов за указанный месяц и год;" + 
                    "\n\t5 - завершить работу программы.");

                switch (Console.ReadKey().KeyChar.ToString()){
                    case "1":
                        FindClientsByWareName();
                        break;
                    case "2":
                        ChangeContactPerson(fname); 
                        break;
                    case "3":
                        FindGoldClient();
                        break;
                    case "4":
                        FindMaxRequiresClient();
                        break;
                    case "5":
                        exit = true;
                        break;
                    default:
                        Console.WriteLine(" Введенный символ не распознан. Попробуйте еще раз.");
                        break;
                }
            }

            Console.WriteLine("\nПрограмма завершилась.Нажмите любую клавишу, чтобы закончить...");
            Console.ReadKey();
        }
    }
}
