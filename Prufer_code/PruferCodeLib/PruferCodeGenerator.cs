using ClosedXML.Excel;
using System.Diagnostics;

namespace PruferCodeLib
{
    public class PruferCodeGenerator
    {
        // Источник трассировки для логирования выполнения
        private static readonly TraceSource trace = new TraceSource("PruferCode");

        public List<(int, int)> ReadEdgesFromExcel(string path)
        {
            trace.TraceInformation("Чтение файла: " + path);

            var edges = new List<(int, int)>(); // Список для хранения рёбер (u, v)

            using (var workbook = new XLWorkbook(path)) // Открываем Excel-файл
            {
                var range = workbook.Worksheet(1).RangeUsed(); // Получаем диапазон с данными

                if (range == null) // Проверка на пустой файл
                    throw new InvalidDataException("Файл пуст");

                foreach (var row in range.Rows()) // Проходим по строкам таблицы
                {
                    if (row.Cell(1).IsEmpty() || row.Cell(2).IsEmpty()) continue; // Пропускаем пустые

                    // Считываем значения двух ячеек и добавляем как ребро
                    edges.Add((row.Cell(1).GetValue<int>(), row.Cell(2).GetValue<int>()));
                }
            }

            return edges;
        }

        public List<int> GeneratePruferCode(List<(int, int)> edges)
        {
            var adj = new Dictionary<int, List<int>>(); // Список смежности

            // Строим список смежности (неориентированное дерево)
            foreach (var pair in edges)
            {
                int u = pair.Item1;
                int v = pair.Item2;

                if (!adj.ContainsKey(u)) adj[u] = new List<int>();
                if (!adj.ContainsKey(v)) adj[v] = new List<int>();

                adj[u].Add(v);
                adj[v].Add(u);
            }

            // Степень каждой вершины — количество соседей
            var degree = new Dictionary<int, int>();
            foreach (var kvp in adj)
                degree[kvp.Key] = kvp.Value.Count;

            // Множество всех листьев (вершин степени 1)
            var leaves = new SortedSet<int>();
            foreach (var kvp in degree)
                if (kvp.Value == 1)
                    leaves.Add(kvp.Key);

            var code = new List<int>(); // Код Прюфера

            // Основной цикл — повторяем n - 2 раз
            for (int i = 0; i < adj.Count - 2; i++)
            {
                int leaf = GetMinFromSet(leaves); // Берём наименьший лист
                leaves.Remove(leaf);

                // Сосед этой вершины (листа)
                int neighbor = adj[leaf].First(v => degree[v] > 0);
                code.Add(neighbor); // Добавляем соседа в код

                // Уменьшаем степень обеих вершин
                degree[leaf]--;
                degree[neighbor]--;

                // Если сосед теперь стал листом — добавляем в множество
                if (degree[neighbor] == 1)
                    leaves.Add(neighbor);
            }

            return code;
        }
        public void WriteCodeToExcel(string path, List<int> code)
        {
            trace.TraceInformation("Запись результата в файл: " + path);


            using (var workbook = new XLWorkbook()) // Создаём новый Excel-файл
            {
                var ws = workbook.Worksheets.Add("PruferCode"); // Добавляем лист
                ws.Cell(1, 1).Value = "Код Прюфера"; // Заголовок

                // Запись значений построчно
                for (int i = 0; i < code.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = code[i];
                }

                ws.Columns().AdjustToContents(); // Автовыравнивание
                workbook.SaveAs(path); // Сохраняем файл
            }
        }

        private int GetMinFromSet(SortedSet<int> set)
        {
            foreach (var value in set)
                return value;
            throw new InvalidOperationException("Множество пусто.");
        }
    }
}

