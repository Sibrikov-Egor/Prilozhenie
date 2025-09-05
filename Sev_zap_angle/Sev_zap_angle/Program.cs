using System.Diagnostics;
using TransportProblem;

namespace Sev_zap_angle
{
    class Program
    {
        static void Main()
        {
            Trace.Listeners.Add(new ConsoleTraceListener());
            Trace.AutoFlush = true;
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";
            try
            {
                double[] supply, demand;
                double[,] costs;
                TransportProblemSolver.ReadDataFromFile(inputFile, out supply, out demand, out costs);
                var solver = new TransportProblemSolver();
                if (!solver.CheckClosed(supply, demand))
                {
                    Console.WriteLine("Ошибка: задача не закрыта");
                    return;
                }
                double totalCost;
                double[,] result = solver.NorthWestCornerMethod(supply, demand, costs, out totalCost);
                solver.WriteResult(result, totalCost, outputFile);
            }
            catch (Exception ex) { Console.WriteLine("Ошибка: " + ex.Message); }
        }
    }
}
