using PruferCodeLib;
using System.Diagnostics;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /// Включение трассировки в консоль
            Trace.Listeners.Add(new ConsoleTraceListener());
            Trace.AutoFlush = true;

            // Пути к входному и выходному файлам
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                var generator = new PruferCodeGenerator(); // Создаём генератор

                // Считываем рёбра из Excel
                var edges = generator.ReadEdgesFromExcel(inputPath);

                if (edges.Count == 0)
                {
                    Console.WriteLine("Файл не содержит рёбер.");
                    return;
                }

                // Генерируем код Прюфера
                var code = generator.GeneratePruferCode(edges);

                // Выводим результат в консоль
                Console.WriteLine("Код Прюфера:");
                Console.WriteLine(string.Join(" ", code));

                // Записываем результат в Excel
                generator.WriteCodeToExcel(outputPath, code);
                Console.WriteLine("Результат записан в файл: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }
        }
    }
}
