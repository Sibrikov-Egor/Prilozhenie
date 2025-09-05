using ClosedXML.Excel;
using System.Diagnostics;

namespace TransportProblem
{
    public class TransportProblemSolver
    {
        private static readonly TraceSource traceSource = new TraceSource("TransportProblemTrace");
        public static void ReadDataFromFile(string filePath, out double[] supply, out double[] demand, out double[,] costs)
        {
            traceSource.TraceEvent(TraceEventType.Information, 0, $"Чтение данных из файла: {filePath}");
            supply = null;
            demand = null;
            costs = null;
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var supplyRow = worksheet.Row(1);
                supply = supplyRow.CellsUsed().Select(c => c.GetValue<double>()).ToArray();
                var demandRow = worksheet.Row(2);
                demand = demandRow.CellsUsed().Select(c => c.GetValue<double>()).ToArray();
                int rows = supply.Length;
                int cols = demand.Length;
                costs = new double[rows, cols];
                for (int i = 0; i < rows; i++)
                {
                    var row = worksheet.Row(i + 3);
                    for (int j = 0; j < cols; j++) costs[i, j] = row.Cell(j + 1).GetValue<double>();
                }
            }
            traceSource.TraceEvent(TraceEventType.Information, 0, "Данные успешно прочитаны");
        }
        public bool CheckClosed(double[] supply, double[] demand)
        {
            double sumSupply = supply.Sum();
            double sumDemand = demand.Sum();
            return Math.Abs(sumSupply - sumDemand) < 0.0001;
        }
        public double[,] NorthWestCornerMethod(double[] supply, double[] demand, double[,] costs, out double totalCost)
        {
            int m = supply.Length, n = demand.Length;
            double[,] result = new double[m, n];
            totalCost = 0;
            double[] tempSupply = (double[])supply.Clone();
            double[] tempDemand = (double[])demand.Clone();
            int i = 0, j = 0;
            while (i < m && j < n)
            {
                double quantity = Math.Min(tempSupply[i], tempDemand[j]);
                result[i, j] = quantity;
                totalCost += quantity * costs[i, j];
                tempSupply[i] -= quantity;
                tempDemand[j] -= quantity;
                if (tempSupply[i] == 0) i++;
                if (tempDemand[j] == 0) j++;
            }
            traceSource.TraceEvent(TraceEventType.Information, 0, $"Метод завершен. Общая стоимость: {totalCost}");
            return result;
        }
        public void WriteResult(double[,] result, double cost, string outputPath)
        {
            traceSource.TraceEvent(TraceEventType.Information, 0, $"Сохранение результатов в {outputPath}");
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Результат");
                for (int i = 0; i < result.GetLength(0); i++)
                {
                    for (int j = 0; j < result.GetLength(1); j++) worksheet.Cell(i + 1, j + 1).Value = result[i, j];
                }
                worksheet.Cell(result.GetLength(0) + 2, 1).Value = $"Общая стоимость: {cost}";
                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(outputPath);
            }
        }
    }
}

