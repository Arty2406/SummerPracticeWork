using Xunit;
using SummerPractice;

namespace SummerPractice.Tests
{
    public class CurrentUserTests
    {
        public CurrentUserTests()
        {
            // Сбрасываем состояние перед каждым тестом, так как класс статический
            CurrentUser.Logout();
        }

        [Fact]
        public void SetUser_ShouldSetPropertiesAndFlags()
        {
            // Act
            CurrentUser.SetUser("TestUser", "Гость");

            // Assert
            Assert.Equal("TestUser", CurrentUser.Login);
            Assert.Equal("Гость", CurrentUser.Role);
            Assert.True(CurrentUser.IsLoggedIn);
            Assert.False(CurrentUser.IsAdmin);
        }

        [Fact]
        public void SetUser_AdminRole_ShouldSetIsAdminTrue()
        {
            // Act
            CurrentUser.SetUser("AdminArty", "Администратор");

            // Assert
            Assert.True(CurrentUser.IsAdmin);
        }

        [Fact]
        public void Logout_ShouldClearAllData()
        {
            // Arrange
            CurrentUser.SetUser("User", "Гость");

            // Act
            CurrentUser.Logout();

            // Assert
            Assert.False(CurrentUser.IsLoggedIn);
            Assert.Null(CurrentUser.Login);
            Assert.Equal("Не авторизован", CurrentUser.GetDisplayName());
        }

        [Fact]
        public void GetDisplayName_WhenLoggedIn_ShouldReturnFormattedString()
        {
            // Arrange
            CurrentUser.SetUser("Ivan", "Гость");

            // Act
            string displayName = CurrentUser.GetDisplayName();

            // Assert
            Assert.Equal("Ivan (Гость)", displayName);
        }
    }
}