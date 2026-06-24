using Xunit;
using SummerPractice;
using System.Reflection;

namespace SummerPractice.Tests
{
    public class DatabaseManagerTests
    {
        [Fact]
        public void VerifyPassword_WithCorrectPassword_ReturnsTrue()
        {
            // Arrange
            string password = "MySecretPass123";

            // Получаем приватный метод через Reflection, чтобы сгенерировать хеш для теста
            var hashMethod = typeof(DatabaseManager).GetMethod("HashPassword",
                BindingFlags.NonPublic | BindingFlags.Static);
            string expectedHash = (string)hashMethod.Invoke(null, new object[] { password });

            // Act
            bool result = DatabaseManager.VerifyPassword(password, expectedHash);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void VerifyPassword_WithWrongPassword_ReturnsFalse()
        {
            // Arrange
            string storedHash = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"; // хеш пустой строки

            // Act
            bool result = DatabaseManager.VerifyPassword("WrongPassword", storedHash);

            // Assert
            Assert.False(result);
        }

        [Fact]
        public void VerifyPassword_WithNullStoredHash_ReturnsFalse()
        {
            // Act & Assert
            Assert.False(DatabaseManager.VerifyPassword("pass", null));
        }
    }
}