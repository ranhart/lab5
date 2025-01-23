using lab5.objects;
using Microsoft.Office.Interop.Excel;
using System;
using System.Threading.Tasks;
using static System.Formats.Asn1.AsnWriter;
using System.Transactions;
using Excel = Microsoft.Office.Interop.Excel;
using static EnterNum;


namespace lab5
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Logger logger = null;
            Database db = null;

            bool validInput = false;
            while (!validInput)
            {
                Console.WriteLine("Как ввести логгирование файлов? (true = добавлять в существующий, false = новый файл): "); // логгер
                string str = Console.ReadLine();
                if (bool.TryParse(str, out bool mode)) 
                {
                    bool logMode = mode;
                    Console.WriteLine("Введите полный путь к лог файлу: ");
                    string logFilePath = Console.ReadLine();
                    logger = new Logger(logFilePath, logMode);
                    validInput = true; 
                }
                else { Console.WriteLine("Ошибка ввода."); }   
            }

            validInput = false;
            while (!validInput)
            {
                Console.WriteLine("Введите полный путь до Excel файла: ");
                string excelPath = Console.ReadLine();
                if (File.Exists(excelPath))
                {
                    
                    db = new Database(excelPath, logger);
                    validInput = true;
                }
                else { Console.WriteLine("Ошибка ввода."); }
            }    

            int task = -1; // выбор задания                                   
            while (task != 0) {
                Console.WriteLine("Выберите задание: " +
                    "\n1 = Чтение базы данных из excel файла." +
                    "\n2 = Просмотр базы данных." +
                    "\n3 = Удаление элементов (по ключу)." +
                    "\n4 = Корректировка элементов (по ключу)." +
                    "\n5 = Добавление элементов." +
                    "\n6 = Сохранить изменения." +
                    "\n0 = Выход.");

                string str = Console.ReadLine();
                if (int.TryParse(str, out task)) {
                    switch (task) { 
                        case 0: break;
                        case 1:
                            db.Read();
                            break;
                        case 2:
                            db.Display();
                            break;
                        case 3:
                            Console.WriteLine("Введите из какой таблицы удалить элемент: 1 - движение товаров, 2 - товары, 3 - категории, 4 - магазины.");
                            Console.WriteLine("Введите ключ эдемента для удаления.");
                            switch (enterNum(1, 4))
                            {
                                case 1: db.DeleteProductMovement(enterNum()); break;
                                case 2: db.DeleteProduct(enterNum()); break;
                                case 3: db.DeleteCategory(enterNum()); break;
                                case 4: db.DeleteShop(Console.ReadLine()); break;
                            }
                            break;
                        case 4:
                            Console.WriteLine("Введите в какой таблице изменить элемент: 1 - движение товаров, 2 - товары, 3 - категории, 4 - магазины.");
                            switch (enterNum(1, 3))
                            {
                                case 1:
                                    Console.WriteLine("Введите id операции, дату, id магазина, артикул, тип операции, количество товара и наличие карты клиента измененного товара:");
                                    ProductMovement productMovement = new ProductMovement(enterNum(),
                                                                                          DateOnly.FromDateTime(DateTime.Parse(Console.ReadLine())),
                                                                                          Console.ReadLine(),
                                                                                          enterNum(),
                                                                                          Console.ReadLine(),
                                                                                          enterNum(),
                                                                                          Console.ReadLine());
                                    Console.WriteLine("Введите артикул старого движения товара: ");
                                    db.ChangeProductMovement(enterNum(), productMovement);
                                    break;
                                case 2:
                                    Console.WriteLine("Введите  артикул, id категории, название предмета, цену при поступлении, цену после скидки, скидку для измененной операции: ");
                                    Product product = new Product(enterNum(),
                                                                  enterNum(),
                                                                  Console.ReadLine(),
                                                                  enterNum(),
                                                                  enterNum(),
                                                                  Console.ReadLine());
                                    Console.WriteLine("Введите артикул старого товара:");
                                    db.ChangeProduct(enterNum(), product);
                                    break;
                                case 3:
                                    Console.WriteLine("Введите id, название и возрастное ограничение изменной категории: ");
                                    Category category = new Category(enterNum(),
                                                                     Console.ReadLine(),
                                                                     Console.ReadLine());
                                    Console.WriteLine("Введите id старой категории:");
                                    db.ChangeCategory(enterNum(), category);
                                    break;
                                case 4:
                                    Console.WriteLine("Введите id, район и адрес изменного магазина: ");
                                    Shop shop = new Shop(Console.ReadLine(),
                                                         Console.ReadLine(),
                                                         Console.ReadLine());
                                    Console.WriteLine("Введите id старого магазина:");
                                    db.ChangeShop(Console.ReadLine(), shop);
                                    break;
                            }
                            break;
                        case 5:
                            Console.WriteLine("Введите в какую таблицу добавить элемент: 1 - движение товаров, 2 - товары, 3 - категории, 4 - магазины.");
                            switch (enterNum(1, 3))
                            {
                                case 1:
                                    Console.WriteLine("Введите id операции, дату, id магазина, артикул, тип операции, количество товара и наличие карты клиента нового товара: ");
                                    ProductMovement productMovement = new ProductMovement(enterNum(),
                                                                                          DateOnly.FromDateTime(DateTime.Parse(Console.ReadLine())),
                                                                                          Console.ReadLine(),
                                                                                          enterNum(),
                                                                                          Console.ReadLine(),
                                                                                          enterNum(),
                                                                                          Console.ReadLine());
                                    db.Add(productMovement);
                                    break;
                                case 2:
                                    Console.WriteLine("Введите  артикул, id категории, название предмета, цену при поступлении, цену после скидки, скидку для новой операции: ");
                                    Product product = new Product(enterNum(),
                                                                  enterNum(),
                                                                  Console.ReadLine(),
                                                                  enterNum(),
                                                                  enterNum(),
                                                                  Console.ReadLine());
                                    db.Add(product);
                                    break;
                                case 3:
                                    Console.WriteLine("Введите id, название и возрастное ограничение новой категории: ");
                                    Category category = new Category(enterNum(),
                                                                     Console.ReadLine(),
                                                                     Console.ReadLine());

                                    db.Add(category);
                                    break;
                                case 4:
                                    Console.WriteLine("Введите id, район и адрес нового магазина: ");
                                    Shop shop = new Shop(Console.ReadLine(),
                                                         Console.ReadLine(),
                                                         Console.ReadLine());
                                    db.Add(shop);
                                    break;
                            }
                            break;
                        case 6: db.Save(); break;
                        default:
                            Console.WriteLine("Ошибка ввода.");
                            break;
                    }           
                }
                else { Console.WriteLine("Ошибка ввода номера задания."); }
            }

        }
    }
}
