using System.Diagnostics;
using TransportProblem;

class Program
{
    static void Main()
    {
        Trace.Listeners.Add(new ConsoleTraceListener());
        Trace.AutoFlush = true;

        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";
        var solver = new TransportProblemSolver();

        if (!solver.ReadInput(inputFile, out var supply, out var demand, out var costs))
        {
            Console.WriteLine("Ошибка чтения данных");
            return;
        }

        if (!solver.IsClosed(supply, demand))
        {
            Console.WriteLine("Ошибка: задача не закрыта");
            return;
        }

        var plan = solver.MinimumCost(supply, demand, costs, out double totalCost);
        solver.SaveToExcel(outputFile, plan, totalCost);

        Console.WriteLine("Решение сохранено в файл: " + outputFile);
    }
}