namespace SummerPractice
{
    /// <summary>
    /// Хранит информацию о текущем залогиненном пользователе.
    /// Доступен из любой формы приложения.
    /// </summary>
    public static class CurrentUser
    {
        /// <summary>Логин текущего пользователя (null, если не залогинен)</summary>
        public static string Login { get; private set; }

        /// <summary>Роль текущего пользователя ("Администратор" или "Гость")</summary>
        public static string Role { get; private set; }

        /// <summary>Залогинен ли пользователь</summary>
        public static bool IsLoggedIn => !string.IsNullOrEmpty(Login);

        /// <summary>Является ли текущий пользователь администратором</summary>
        public static bool IsAdmin => Role == "Администратор";

        /// <summary>
        /// Устанавливает данные текущего пользователя (вызывается при успешном логине).
        /// </summary>
        public static void SetUser(string login, string role)
        {
            Login = login;
            Role = role;
        }

        /// <summary>
        /// Сбрасывает данные пользователя (при выходе из системы).
        /// </summary>
        public static void Logout()
        {
            Login = null;
            Role = null;
        }

        /// <summary>
        /// Возвращает строку для отображения, например "admin (Администратор)".
        /// </summary>
        public static string GetDisplayName()
        {
            if (!IsLoggedIn) return "Не авторизован";
            return $"{Login} ({Role})";
        }
    }
}