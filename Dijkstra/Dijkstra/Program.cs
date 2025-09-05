using DijkstraAlgorithmLib;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        Trace.Listeners.Add(new ConsoleTraceListener());
        Trace.AutoFlush = true;

        string input = "input.xlsx";
        string output = "output.xlsx";
        int startVertex = 0;

        try
        {
            var graph = Dijkstra.ReadGraph(input);
            var result = Dijkstra.FindPaths(graph, startVertex);
            Dijkstra.WriteResult(output, result);

            Console.WriteLine("Кратчайшие расстояния:");
            for (int i = 0; i < result.Length; i++)
            {
                if (result[i] == int.MaxValue)
                    Console.WriteLine($"До вершины {i}: недостижима");
                else
                    Console.WriteLine($"До вершины {i}: {result[i]}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }
}
