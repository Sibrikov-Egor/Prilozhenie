using ClosedXML.Excel;
using System.Diagnostics;

namespace TransportProblem
{
    public class TransportProblemSolver
    {
        private static readonly TraceSource trace = new TraceSource("TransportProblemTrace");

        public bool ReadInput(string filePath, out double[] supply, out double[] demand, out double[,] costs)
        {
            supply = null;
            demand = null;
            costs = null;

            try
            {
                var wb = new XLWorkbook(filePath);
                var ws = wb.Worksheet(1);

                int supplyCount = (int)ws.Cell(1, 1).GetDouble();
                int demandCount = (int)ws.Cell(1, 2).GetDouble();

                supply = new double[supplyCount];
                demand = new double[demandCount];
                costs = new double[supplyCount, demandCount];

                for (int i = 0; i < supplyCount; i++)
                    supply[i] = ws.Cell(2, i + 1).GetDouble();

                for (int j = 0; j < demandCount; j++)
                    demand[j] = ws.Cell(3, j + 1).GetDouble();

                for (int i = 0; i < supplyCount; i++)
                    for (int j = 0; j < demandCount; j++)
                        costs[i, j] = ws.Cell(4 + i, j + 1).GetDouble();

                trace.TraceEvent(TraceEventType.Information, 0, "Данные успешно считаны из Excel");
                return true;
            }
            catch (Exception ex)
            {
                trace.TraceEvent(TraceEventType.Error, 0, ex.Message);
                return false;
            }
        }

        public bool IsClosed(double[] supply, double[] demand)
        {
            return Math.Abs(supply.Sum() - demand.Sum()) < 0.0001;
        }

        public double[,] MinimumCost(double[] supply, double[] demand, double[,] cost, out double totalCost)
        {
            int m = supply.Length, n = demand.Length;
            double[,] result = new double[m, n];
            totalCost = 0;
            double[] s = (double[])supply.Clone();
            double[] d = (double[])demand.Clone();
            bool[,] used = new bool[m, n];

            while (true)
            {
                int minI = -1, minJ = -1;
                double min = double.MaxValue;
                for (int i = 0; i < m; i++)
                    for (int j = 0; j < n; j++)
                        if (!used[i, j] && cost[i, j] < min)
                        {
                            min = cost[i, j];
                            minI = i;
                            minJ = j;
                        }

                if (minI == -1) break;

                double q = Math.Min(s[minI], d[minJ]);
                result[minI, minJ] = q;
                totalCost += q * cost[minI, minJ];
                s[minI] -= q;
                d[minJ] -= q;

                if (s[minI] == 0)
                    for (int k = 0; k < n; k++) used[minI, k] = true;
                if (d[minJ] == 0)
                    for (int k = 0; k < m; k++) used[k, minJ] = true;
            }

            trace.TraceEvent(TraceEventType.Information, 0, "Решение методом минимального элемента завершено");
            return result;
        }

        public void SaveToExcel(string path, double[,] plan, double totalCost)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Результат");

            for (int i = 0; i < plan.GetLength(0); i++)
                for (int j = 0; j < plan.GetLength(1); j++)
                    ws.Cell(i + 1, j + 1).Value = plan[i, j];

            ws.Cell(plan.GetLength(0) + 2, 1).Value = "Общая стоимость: ";
            ws.Cell(plan.GetLength(0) + 2, 2).Value = totalCost;
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }


    }
}

