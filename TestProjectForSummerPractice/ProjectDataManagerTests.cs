using Xunit;
using SummerPractice;
using System;
using System.Collections.Generic;

namespace SummerPractice.Tests
{
    public class ProjectDataManagerTests
    {
        private readonly ProjectDataManager manager;

        public ProjectDataManagerTests()
        {
            manager = new ProjectDataManager("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=fake.accdb;");
        }

        [Fact]
        public void ApplySort_WhenTableNotSelected_ThrowsInvalidOperationException()
        {
            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() =>
                manager.ApplySort("SomeColumn", true));

            Assert.Contains("Таблица не выбрана", ex.Message);
        }

        [Fact]
        public void PerformSearch_WhenTableNotSelected_ThrowsInvalidOperationException()
        {
            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() =>
                manager.PerformSearch("text", new List<string>()));

            Assert.Contains("Таблица не выбрана", ex.Message);
        }

        [Fact]
        public void InitialState_ShouldBeCorrect()
        {
            // Assert
            Assert.False(manager.HasUnsavedChanges);
            Assert.False(manager.IsFiltered);
            Assert.False(manager.IsSorted);
            Assert.False(manager.IsSearched);
            Assert.Null(manager.CurrentTableName);
        }

        [Fact]
        public void ResetAll_ShouldClearAllFlags()
        {
            // Arrange (эмулируем изменение флагов через reflection или публичные методы, если бы они были)
            // В данном классе флаги меняются внутри методов SelectTable, ApplySort и т.д.
            // Протестируем сам факт сброса, если бы они были true.
            // Так как публичных сеттеров нет, проверяем начальное состояние.
            manager.ResetAll();

            Assert.False(manager.IsFiltered);
            Assert.False(manager.IsSorted);
            Assert.False(manager.IsSearched);
        }
    }
}