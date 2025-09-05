using ClosedXML.Excel;
using System.Diagnostics;

namespace DijkstraAlgorithmLib
{
    public class Dijkstra
    {
        private static readonly TraceSource traceSource = new TraceSource("DijkstraTrace");
        public static int[,] ReadGraph(string filePath)
        {
            traceSource.TraceEvent(TraceEventType.Information, 0, $"Чтение графа из файла: {filePath}");

            var workbook = new XLWorkbook(filePath);
            var sheet = workbook.Worksheet(1);
            var range = sheet.RangeUsed();
            int size = range.RowCount();
            int[,] graph = new int[size, size];

            for (int i = 0; i < size; i++)
                for (int j = 0; j < size; j++)
                    graph[i, j] = sheet.Cell(i + 1, j + 1).GetValue<int>();

            traceSource.TraceEvent(TraceEventType.Information, 0, $"Прочитано {size} вершин");
            return graph;
        }

        public static int[] FindPaths(int[,] graph, int start)
        {
            int n = graph.GetLength(0);
            int[] dist = new int[n];
            bool[] visited = new bool[n];

            for (int i = 0; i < n; i++) dist[i] = int.MaxValue;
            dist[start] = 0;

            for (int count = 0; count < n - 1; count++)
            {
                int u = -1, min = int.MaxValue;
                for (int i = 0; i < n; i++)
                    if (!visited[i] && dist[i] < min) { min = dist[i]; u = i; }

                if (u == -1) break;
                visited[u] = true;
                traceSource.TraceEvent(TraceEventType.Verbose, 0, $"Обработка вершины {u}");

                for (int v = 0; v < n; v++)
                    if (!visited[v] && graph[u, v] > 0 && dist[u] + graph[u, v] < dist[v])
                    {
                        dist[v] = dist[u] + graph[u, v];
                        traceSource.TraceEvent(TraceEventType.Verbose, 0, $"Обновлено расстояние до {v}: {dist[v]}");
                    }
            }

            traceSource.TraceEvent(TraceEventType.Information, 0, "Расчет завершен");
            return dist;
        }

        public static void WriteResult(string path, int[] distances)
        {
            traceSource.TraceEvent(TraceEventType.Information, 0, $"Запись результата в файл: {path}");

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Результат");
            ws.Cell(1, 1).Value = "Вершина";
            ws.Cell(1, 2).Value = "Расстояние";
            for (int i = 0; i < distances.Length; i++)
            {
                ws.Cell(i + 2, 1).Value = i;
                if (distances[i] == int.MaxValue)
                    ws.Cell(i + 2, 2).Value = "недостижима";
                else
                    ws.Cell(i + 2, 2).Value = distances[i];
            }
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }
    }
}

