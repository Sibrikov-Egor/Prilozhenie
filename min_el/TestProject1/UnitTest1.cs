using System;
using System.IO;
using ClosedXML.Excel;
using Xunit;

namespace TestProject1
{
    public class TransportProblemSolverTests : IDisposable
    {
        private readonly TransportProblem.TransportProblemSolver _solver;
        private readonly string _testInputPath = "test_input.xlsx";
        private readonly string _testOutputPath = "test_output.xlsx";

        public TransportProblemSolverTests()
        {
            _solver = new TransportProblem.TransportProblemSolver();
            CreateTestExcelFile();
        }

        public void Dispose()
        {
            if (File.Exists(_testInputPath))
                File.Delete(_testInputPath);
            if (File.Exists(_testOutputPath))
                File.Delete(_testOutputPath);
        }

        private void CreateTestExcelFile()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                // Количество поставщиков и потребителей
                worksheet.Cell(1, 1).Value = 2; // Поставщики
                worksheet.Cell(1, 2).Value = 3; // Потребители

                // Запасы
                worksheet.Cell(2, 1).Value = 50;
                worksheet.Cell(2, 2).Value = 30;

                // Потребности
                worksheet.Cell(3, 1).Value = 20;
                worksheet.Cell(3, 2).Value = 40;
                worksheet.Cell(3, 3).Value = 20;

                // Матрица стоимостей
                worksheet.Cell(4, 1).Value = 3; worksheet.Cell(4, 2).Value = 5; worksheet.Cell(4, 3).Value = 7;
                worksheet.Cell(5, 1).Value = 2; worksheet.Cell(5, 2).Value = 4; worksheet.Cell(5, 3).Value = 6;

                workbook.SaveAs(_testInputPath);
            }
        }

        // Проверяет чтение данных из Excel: возврат true, корректность запасов, потребностей и стоимостей
        [Fact]
        public void ReadInput_ValidFile_ReturnsTrueAndCorrectData()
        {
            // Act
            bool result = _solver.ReadInput(_testInputPath,
                out double[] supply, out double[] demand, out double[,] costs);

            // Assert
            Assert.True(result);
            Assert.Equal(new double[] { 50, 30 }, supply);
            Assert.Equal(new double[] { 20, 40, 20 }, demand);
            Assert.Equal(3, costs[0, 0]);
            Assert.Equal(6, costs[1, 2]);
        }

        // Проверяет обработку отсутствия файла: возврат false и null выходных параметров
        [Fact]
        public void ReadInput_InvalidFile_ReturnsFalse()
        {
            // Act
            bool result = _solver.ReadInput("nonexistent_file.xlsx",
                out double[] supply, out double[] demand, out double[,] costs);

            // Assert
            Assert.False(result);
            Assert.Null(supply);
            Assert.Null(demand);
            Assert.Null(costs);
        }


        // Проверяет определение закрытости задачи при равных суммах запасов и потребностей
        [Fact]
        public void IsClosed_WhenSumsEqual_ReturnsTrue()
        {
            // Arrange
            double[] supply = { 50, 30 };
            double[] demand = { 20, 40, 20 };

            // Act
            bool result = _solver.IsClosed(supply, demand);

            // Assert
            Assert.True(result);
        }

        // Проверяет определение незакрытости задачи при неравных суммах
        [Fact]
        public void IsClosed_WhenSumsNotEqual_ReturnsFalse()
        {
            // Arrange
            double[] supply = { 50, 30 };
            double[] demand = { 20, 40, 25 };

            // Act
            bool result = _solver.IsClosed(supply, demand);

            // Assert
            Assert.False(result);
        }
    }
}