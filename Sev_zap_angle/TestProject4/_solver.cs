using NUnit.Framework;
using System;
using System.IO;
using ClosedXML.Excel;
using TransportProblem;

namespace TestProject4
{
    public class Tests
    {
        private TransportProblemSolver _solver;
        private string _testInputPath = "test_input.xlsx";
        private string _testOutputPath = "test_output.xlsx";

        [Test]
        public void ReadDataFromFile_ValidFile_ReadsCorrectData()
        {
            // Act
            TransportProblemSolver.ReadDataFromFile(_testInputPath,
                out double[] supply, out double[] demand, out double[,] costs);

            // Assert
            Assert.AreEqual(new double[] { 30, 20 }, supply);
            Assert.AreEqual(new double[] { 10, 25, 15 }, demand);
            Assert.AreEqual(2, costs.GetLength(0));
            Assert.AreEqual(3, costs.GetLength(1));
            Assert.AreEqual(2, costs[0, 0]);
            Assert.AreEqual(8, costs[1, 2]);
        }

        [Test]
        public void CheckClosed_WhenSumsEqual_ReturnsTrue()
        {
            // Arrange
            var supply = new double[] { 50, 30 };
            var demand = new double[] { 40, 20, 20 };

            // Act
            var result = _solver.CheckClosed(supply, demand);

            // Assert
            Assert.IsTrue(result);
        }

        [Test]
        public void CheckClosed_WhenSumsNotEqual_ReturnsFalse()
        {
            // Arrange
            var supply = new double[] { 50, 30 };
            var demand = new double[] { 40, 20, 25 };

            // Act
            var result = _solver.CheckClosed(supply, demand);

            // Assert
            Assert.IsFalse(result);
        }
    }
}
