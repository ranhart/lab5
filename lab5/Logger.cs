using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class Logger
    {
        private string logFilePath;
        private bool logMode;

        public Logger(string path, bool mode) { // конструктор логгера
            logFilePath = path;
            if (!File.Exists(logFilePath)) {
                File.WriteAllText(logFilePath, "Лог файл создан." + DateTime.Now + Environment.NewLine );
            }
            else if (!logMode)
            {
                File.WriteAllText(logFilePath, "Лог-файл очищен: " + DateTime.Now + Environment.NewLine);
            }
        }
        public void Log(string message) // запись в лог файл
        {
            File.AppendAllText(logFilePath, $"{DateTime.Now}: {message}");
        }
    }
}
