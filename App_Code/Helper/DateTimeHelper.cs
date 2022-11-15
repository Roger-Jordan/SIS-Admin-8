using System;


namespace SIS.Helper
{
    public abstract class DateTimeHelper
    {
        public static string ToGermanDateString(DateTime date)
        {
            if (date == null)
                return null;

            string result = date.ToString("dd.MM.yyyy");
            return result;
        }
    }
}
