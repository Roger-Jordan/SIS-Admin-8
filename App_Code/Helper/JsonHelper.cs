using System.Text.RegularExpressions;


namespace SIS.Helper
{
    public abstract class JsonHelper
    {
        public static string GetErrorWithErrorMessage(string errorMessage, bool escapeDoubleQuotes = true)
        {
            if (escapeDoubleQuotes)
                errorMessage = EscapeDoubleQuotes(errorMessage);

            string json = string.Format("{{\"status\":\"error\",\"message\":\"{0}\"}}", errorMessage);
            return json;
        }

        public static string GetSuccessWithRedirectUrl(string redirectUrl)
        {
            string json = "{\"status\":\"success\",\"redirectUrl\":\"" + redirectUrl + "\"}";
            return json;
        }

        public static string GetSuccessWithAutoSubmit()
        {
            string json = "{\"status\":\"success\",\"autoSubmit\":\"true\"}";
            return json;
        }

        private static string EscapeDoubleQuotes(string value)
        {
            return Regex.Replace(value, "(?<!\\\\)\"", "\\\"");
        }
    }
}