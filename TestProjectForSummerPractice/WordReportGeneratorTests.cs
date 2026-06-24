using Xunit;
using SummerPractice;
using System.Windows.Forms;

namespace SummerPractice.Tests
{
    public class WordReportGeneratorTests
    {
        [Fact]
        public void CreateReport_WithNullDataGridView_ReturnsError()
        {
            // Act
            var result = WordReportGenerator.CreateReport(null, "Table1", "User");

            // Assert
            Assert.False(result.Success);
            Assert.Contains("Нет данных", result.ErrorMessage);
        }

        [Fact]
        public void CreateReport_WithNullDataSource_ReturnsError()
        {
            // Arrange
            var dgv = new DataGridView { DataSource = null };

            // Act
            var result = WordReportGenerator.CreateReport(dgv, "Table1", "User");

            // Assert
            Assert.False(result.Success);
            Assert.Contains("Нет данных", result.ErrorMessage);
        }
    }
}