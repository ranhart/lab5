using System;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace lab5
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Как ввести логгирование файлов? (1 = тот же файл, 0 = новый файл): "); // логгер
            string str = Console.ReadLine();
            if (bool.TryParse(str, out bool mode)) 
            {
                bool logMode = mode;
                Console.WriteLine("Введите путь к лог файлу: ");
                string logFilePath = Console.ReadLine();
                Logger logger = new Logger(logFilePath, logMode);
            }
            else { Console.WriteLine("Ошибка ввода."); }     

            int task = -1; // выбор задания                                   
            while (task != 0) {
                Console.WriteLine("Выберите задание: " +
                    "\n1 = Чтение базы данных из excel файла." +
                    "\n2 = Просмотр базы данных." +
                    "\n3 = Удаление элементов (по ключу)." +
                    "\n4 = Корректировка элементов (по ключу)." +
                    "\n5 = Добавление элементов." +
                    "\n6 = Реализация 4 запросов." +
                    "\n7 = логгирование" +
                    "\n0 = Выход.");

                str = Console.ReadLine();
                if (int.TryParse(str, out task)) {
                    switch (task) {                    
                        case 1:


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
