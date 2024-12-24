using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class Logger
    {
        private string _logFilePath;

        public Logger(string logFilePath, bool logMode) { // конструктор логгера
           
            if (logMode)
            {
                if (File.Exists(logFilePath))
                {
                    _logFilePath = logFilePath;
                }
                else { throw new FileNotFoundException("Файл отсутствует"); }
            }
            else
            {
                _logFilePath = logFilePath;
                File.Create(logFilePath).Close();
                File.WriteAllText(logFilePath, "Лог файл создан. " + DateTime.Now + Environment.NewLine);
            }
        }
        public void Log(string message) // запись в лог файл
        {
            File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}\n");
        }
    }
}
