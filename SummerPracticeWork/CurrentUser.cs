namespace SummerPractice
{
    public static class CurrentUser
    {
        public static string Login { get; private set; }

        public static string Role { get; private set; }

        public static bool IsLoggedIn => !string.IsNullOrEmpty(Login);

        public static bool IsAdmin => Role == "Администратор";

        public static void SetUser(string login, string role)
        {
            Login = login;
            Role = role;
        }

        public static void Logout()
        {
            Login = null;
            Role = null;
        }

        public static string GetDisplayName()
        {
            if (!IsLoggedIn) return "Не авторизован";
            return $"{Login} ({Role})";
        }
    }
}